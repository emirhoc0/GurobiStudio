// Real UI (mock sheet) -> real helper -> real Gurobi. Starts its own helper on a spare port with a
// throwaway data folder, so it never touches your approvals or run history. Needs `npm run build`.
import { spawn } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { webkit } from "playwright";

const PORT = 8799;
const ROOT = new URL("../../", import.meta.url).pathname;
const SHOTS = process.env.GS_SHOTS ?? join(tmpdir(), "gs-e2e-shots");
mkdirSync(SHOTS, { recursive: true });
const data = mkdtempSync(join(tmpdir(), "gs-e2e-"));
const helper = spawn(join(ROOT, "helper/.venv/bin/python"), ["-m", "gurobistudio"], {
  cwd: join(ROOT, "helper"),
  env: { ...process.env, GS_PORT: String(PORT), GS_DATA_DIR: data },
  stdio: "inherit",
});
process.on("exit", () => helper.kill()); // even when a step throws, or the next run finds the port taken

let failures = 0;
const check = (name, ok, detail = "") => {
  console.log(`${ok ? "PASS" : "FAIL"}  ${name}${ok ? "" : `  -> ${detail}`}`);
  if (!ok) failures++;
};
const HEADER = "from SolverStudio import *\nfrom gurobipy import *\n";
let unique = 0;

async function waitForHelper() {
  process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";
  for (let i = 0; i < 100; i++) {
    try {
      await fetch(`https://localhost:${PORT}/`);
      return;
    } catch {
      await new Promise((r) => setTimeout(r, 100));
    }
  }
  throw new Error("The helper did not start.");
}

await waitForHelper();
const key = readFileSync(join(data, "page_key"), "utf8").trim();
const URL_ = `https://localhost:${PORT}/?k=${key}`;
const browser = await webkit.launch();
const context = await browser.newContext({
  ignoreHTTPSErrors: true,
  viewport: { width: 380, height: 900 },
});
const page = await context.newPage();
const pageErrors = [];
page.on("pageerror", (e) => pageErrors.push(e.message));

const openPane = async () => {
  await page.goto(URL_);
  await page.getByRole("tab", { name: "Model" }).waitFor();
  await page.locator(".cm-content").waitFor();
};
const setModel = (text) =>
  page.evaluate((t) => window.gsModels.set({ ...window.gsModels.get(), "mock-sheet": t }), text);
const setItems = (items) =>
  page.evaluate((i) => window.gsItems.set({ ...window.gsItems.get(), "mock-sheet": i }), items);
const setCells = (cells) =>
  page.evaluate((c) => Object.entries(c).forEach(([k, v]) => window.mockCells.set(k, v)), cells);
const cell = (address) => page.evaluate((a) => window.mockCells.get(a), address);
const item = (name, type, cellRange, indexRanges = [], readOnly = false) => ({
  name,
  type,
  cellRange,
  indexRanges,
  readOnly,
});
const dialog = () => page.getByRole("alertdialog");
const log = () => page.getByTestId("log").innerText();
const output = () => page.getByRole("tab", { name: "Output" }).click();
const solveButton = () => page.getByRole("button", { name: "Solve Model" });

/** Clicks Solve, approves the trust prompt, waits for the run to end. Returns the chip's text. */
async function solveAndApprove() {
  await output();
  await solveButton().click();
  await dialog().getByRole("button", { name: "Run" }).click();
  return finished();
}
async function finished() {
  await page.waitForFunction(
    () => !document.body.innerText.match(/\bRunning\b/) && document.querySelector('[data-testid="log"]'),
    null,
    { timeout: 120000 },
  );
  return page.locator(".fui-Badge").last().innerText();
}

// ---- Editor
await openPane();
check("editor shows the template", (await page.locator(".cm-content").innerText()).includes("from SolverStudio import *"));
await page.locator(".cm-content").click();
await page.keyboard.press("Meta+ArrowDown");
await page.keyboard.insertText("x = 1  # typed");
check(
  "typing reaches the model store",
  (await page.evaluate(() => window.gsModels.get()["mock-sheet"])).includes("x = 1  # typed"),
);
await page.getByRole("tab", { name: "Output" }).click();
await page.getByRole("tab", { name: "Model" }).click();
await page.locator(".cm-content").click();
await page.keyboard.press("Meta+z");
check(
  "undo survives a tab switch",
  !(await page.evaluate(() => window.gsModels.get()["mock-sheet"])).includes("# typed"),
);
await page.keyboard.press("Meta+Shift+z");

// ---- Trust prompt
const trusted = `${HEADER}print("trust ${++unique}")\n`;
await setModel(trusted);
await output();
await solveButton().click();
await dialog().waitFor();
check("first run shows the code", (await page.getByTestId("confirm-code").innerText()).includes(`trust ${unique}`));
await dialog().getByRole("button", { name: "Cancel" }).click();
await page.getByText("Click Solve Model to run").waitFor();
check("cancel runs nothing", true);
await solveButton().click();
await dialog().getByRole("button", { name: "Run" }).click();
await finished();
check("approved model ran", (await log()).includes(`trust ${unique}`));
await solveButton().click();
await page.waitForTimeout(300);
check("same code runs without a prompt", !(await dialog().isVisible()));
await finished();
await setModel(`${trusted}print("changed")\n`);
await solveButton().click();
await dialog().waitFor();
check("changed code prompts again", true);
await dialog().getByRole("button", { name: "Cancel" }).click();

// ---- Error card and jump to line
await setModel(`${HEADER}x = 1\ny = 1 / 0\n`);
const chip = await solveAndApprove();
check("error chip", chip === "Error", chip);
await page.getByRole("button", { name: 'File "model.py", line 4' }).waitFor();
await page.getByRole("button", { name: "Go to line 4" }).click();
await page.locator(".cm-content").waitFor();
const selected = await page.evaluate(() => getSelection().toString());
check("jump selects the failing line", selected.includes("y = 1 / 0"), selected);
await page.screenshot({ path: join(SHOTS, "jump.png") });

await setModel(`${HEADER}print 'old style'\n`);
await solveAndApprove();
check("Python 2 hint", await page.getByText("Add parentheses").isVisible());

// ---- Infeasible (acceptance 2)
await setCells({ K1: "untouched", K2: 5 });
await setItems([item("note", "single", "K1"), item("flag", "single", "K2")]);
await setModel(`${HEADER}m = Model()\nv = m.addVar(ub=1)\nm.addConstr(v >= 2)\nm.optimize()\nflag = 9\n`);
const inf = await solveAndApprove();
check("infeasible chip", inf === "Infeasible", inf);
check("infeasible writes only what was assigned", (await cell("K1")) === "untouched" && (await cell("K2")) === 9);

// ---- Formula warning
await setCells({ L1: 1, L2: 2 });
await page.evaluate(() => window.mockExcel.formulas.set("L1", "=1+0"));
await setItems([item("a", "single", "L1"), item("b", "single", "L2")]);
await setModel(`${HEADER}a = 10\nb = 20\n`);
await output();
await solveButton().click();
await dialog().getByRole("button", { name: "Run" }).click();
await dialog().getByText("L1").waitFor();
check("formula prompt names the cell", true);
await dialog().getByRole("button", { name: "Keep formulas" }).click();
await finished();
check("keep leaves the formula and writes the rest", (await cell("L1")) === 1 && (await cell("L2")) === 20);
await page.getByText("Kept 1 formula cell").waitFor();
await solveButton().click();
await dialog().getByRole("button", { name: "Overwrite" }).click();
await finished();
check("overwrite replaces the formula", (await cell("L1")) === 10);
await page.evaluate(() => window.gsItems.set({}));

// ---- Stop mid-solve, responsiveness, shortcut (acceptance 3)
const LONG = `${HEADER}import random
random.seed(1)
m = Model()
n = 400
x = m.addVars(n, vtype=GRB.BINARY, obj=[random.randint(1, 99) for _ in range(n)])
for _ in range(600):
    m.addConstr(sum(x[random.randrange(n)] for _ in range(12)) >= 2)
m.params.TimeLimit = 120
m.optimize()
`;
await setModel(LONG);
await output();
await solveButton().click();
await dialog().getByRole("button", { name: "Run" }).click();
await page.waitForFunction(() => document.querySelector('[data-testid="log"]')?.innerText.includes("Nodes"), null, { timeout: 60000 });
const before = (await log()).split("\n").length;
await page.waitForTimeout(2500);
check("log streams while solving", (await log()).split("\n").length > before);
const t0 = Date.now();
await page.getByRole("tab", { name: "Settings" }).click();
await page.getByRole("tab", { name: "Output" }).click();
check("pane responds during a solve", Date.now() - t0 < 1500, `${Date.now() - t0} ms`);
const stopAt = Date.now();
await page.keyboard.press("Meta+.");
const stopped = await finished();
check("⌘. stops it", stopped === "Stopped", stopped);
check("stops within 5 s", Date.now() - stopAt < 5000, `${Date.now() - stopAt} ms`);

await setModel(`${HEADER}print("from shortcut")\n`);
await page.getByRole("tab", { name: "Model" }).click();
await page.locator(".cm-content").click();
await page.keyboard.press("Meta+Enter");
await dialog().getByRole("button", { name: "Run" }).click();
await finished();
check("⌘↵ solves from the editor and shows Output", (await log()).includes("from shortcut"));

// ---- Copy and clear
await page.getByRole("button", { name: "Clear" }).click();
await page.getByText("Click Solve Model to run").waitFor();
check("Clear empties the log", true);

// ---- Pop out
await page.getByRole("tab", { name: "Model" }).click();
const [popup] = await Promise.all([
  context.waitForEvent("page"),
  page.getByRole("button", { name: "Pop out" }).click(),
]);
await popup.locator(".cm-content").waitFor();
check("pop-out shows the model", (await popup.locator(".cm-content").innerText()).includes("from shortcut"));
await popup.locator(".cm-content").click();
await popup.keyboard.press("Meta+ArrowDown");
await popup.keyboard.insertText("# from the dialog");
await page.waitForFunction(() => window.gsModels.get()["mock-sheet"].includes("# from the dialog"));
check("dialog edits reach the pane", true);
await setModel(`${HEADER}# from the pane\n`);
await popup.getByText("# from the pane").waitFor();
check("pane edits reach the dialog", true);
const big = `${HEADER}${"# padding line\n".repeat(20000)}`;
await setModel(big);
await popup.getByText("padding", { exact: false }).first().waitFor();
check("a 300 KB model reaches the dialog", (await popup.evaluate(() => document.querySelector(".cm-gutters").innerText.length)) > 0);
await setModel(`${HEADER}print("solved from dialog")\n`);
await popup.getByText("solved from dialog").waitFor();
await popup.locator(".cm-content").click();
await popup.keyboard.press("Meta+Enter");
await dialog().getByRole("button", { name: "Run" }).click();
await finished();
check("⌘↵ in the dialog solves in the pane", (await log()).includes("solved from dialog"));
await popup.close();

// ---- Settings reach the run
await page.evaluate(() => window.gsSettings.set({ ...window.gsSettings.get(), timeLimit: "5" }));
await page.getByRole("tab", { name: "Settings" }).click();
await page.getByLabel("TimeLimit").waitFor();
check("settings field shows the value", (await page.getByLabel("TimeLimit").inputValue()) === "5");
await page.getByRole("button", { name: "Add parameter" }).click();
await page.getByLabel("Parameter 1 name").fill("MIPFocus");
await page.getByLabel("Parameter 1 value").fill("1");
await setModel(`${HEADER}m = Model()\nm.optimize()\nprint("focus", m.Params.MIPFocus, m.Params.TimeLimit)\n`);
await solveAndApprove();
const paramLog = await log();
check("parameters are applied before the model", /focus 1 5(\.0)?\b/.test(paramLog), paramLog);
await page.getByRole("tab", { name: "Settings" }).click();
await page.getByLabel("Parameter 1 name").fill("NoSuchParam");
await output();
await solveButton().click();
const badChip = await finished();
check("a bad parameter gets an error card", badChip === "Error" && (await page.getByText("NoSuchParam").first().isVisible()), badChip);
await page.getByRole("tab", { name: "Settings" }).click();
await page.getByRole("button", { name: "Remove parameter 1" }).click();

// ---- Check License (acceptance 8)
await page.getByRole("button", { name: "Check License" }).click();
await page.getByTestId("license-report").waitFor({ timeout: 90000 });
const report = await page.getByTestId("license-report").innerText();
check("Check License reports the Academic license", report.includes("Academic") && !report.includes("size-limited"), report);
check("license expiry is a real date or absent", !report.includes("9999"), report);
await page.screenshot({ path: join(SHOTS, "settings.png"), fullPage: true });
const overflow = await page.evaluate(() => {
  const pane = document.querySelector('[role="tablist"]').nextElementSibling;
  return [pane.scrollWidth, pane.clientWidth];
});
check("Settings fits the pane", overflow[0] <= overflow[1], overflow.join(" > "));

// ---- Helper settings
await page.getByLabel("Keep runs").fill("7");
await page.getByRole("button", { name: "Save helper settings" }).click();
await page.getByText("Helper settings saved").first().waitFor();
await page.getByLabel("Python path").fill("/no/such/python");
await page.getByRole("button", { name: "Save helper settings" }).click();
await page.getByText("is not an executable").first().waitFor();
check("helper settings save and refuse a bad path", true);
await page.getByLabel("Python path").fill("");

// ---- Save and reopen (acceptance 5; a reload is the mock's reopen)
const tricky = 'a = 1\tb\r\nprint("<x> & ]]> çğüş ✓")\n';
await setModel(tricky);
await setItems([item("kept", "indexed", "B1:B2", ["A1:A2"], true)]);
await page.evaluate(() => window.gsSettings.set({ ...window.gsSettings.get(), mipGap: "0.05", outputFlag: false }));
await page.waitForTimeout(600);
await openPane();
const after = await page.evaluate(() => ({
  model: window.gsModels.get()["mock-sheet"],
  items: window.gsItems.get()["mock-sheet"],
  settings: window.gsSettings.get(),
}));
check("model text restored exactly", after.model === tricky, JSON.stringify(after.model));
check("items restored", after.items?.[0]?.name === "kept" && after.items[0].readOnly === true);
check("settings restored", after.settings.mipGap === "0.05" && after.settings.outputFlag === false && after.settings.timeLimit === "5");
await page.waitForTimeout(300);
check("nothing runs on open", await page.getByRole("tab", { name: "Model", selected: true }).isVisible());

// A trusted model stays trusted across a reopen, since the workbook id is stored.
await setItems([]);
await page.evaluate(() => window.gsSettings.set({ ...window.gsSettings.get(), outputFlag: true }));
await setModel(`${HEADER}print("still trusted")\n`);
await solveAndApprove();
await page.waitForTimeout(600);
await openPane();
await setModel(`${HEADER}print("still trusted")\n`);
await output();
await solveButton().click();
await page.waitForTimeout(400);
check("approval survives a reopen", !(await dialog().isVisible()));
await finished();

// ---- Unreadable state is never overwritten
const garbage = '{"schema":99,"workbookId":"x"}';
await page.evaluate((g) => localStorage.setItem("gs-mock-state", g), garbage);
await openPane();
await page.getByRole("alert").getByText("could not be read").waitFor();
await setModel("# edited\n");
await page.waitForTimeout(600);
check("a newer schema is left untouched", (await page.evaluate(() => localStorage.getItem("gs-mock-state"))) === garbage);
await page.evaluate(() => localStorage.removeItem("gs-mock-state"));
await openPane();

// ---- Examples (acceptance 1): every one inserts, prompts, solves and matches brute force
const examples = Object.fromEntries(
  ["diet", "transportation", "knapsack", "assignment"].map((n) => [
    n,
    JSON.parse(readFileSync(join(ROOT, "examples", `${n}.json`), "utf8")),
  ]),
);
const close = (a, b) => typeof a === "number" && Math.abs(a - b) < 1e-6;
for (const example of Object.values(examples)) {
  const { name, expected } = example;
  await page.getByRole("tab", { name: "Model" }).click();
  await page.getByRole("button", { name: "File" }).click();
  await page.getByRole("menuitem", { name: "View Examples" }).click();
  if (name === "Diet") await page.screenshot({ path: join(SHOTS, "examples-dialog.png") });
  await page.getByRole("button", { name: `Insert ${name}` }).click();
  await page.getByText(`Model for ${name}`).waitFor();
  check(`${name}: model shows in the editor`, (await page.locator(".cm-content").innerText()).includes(`Model("${name.toLowerCase()}")`));
  const chip = await solveAndApprove();
  check(`${name}: solves to Optimal`, chip === "Optimal", chip);
  check(`${name}: objective in the log`, /Optimal (objective|solution found)/.test(await log()));
  const output_ = await log();
  check(`${name}: changed items are starred`, /Data items: .*\*/.test(output_), output_);
  const wrong = [];
  for (const [address, want] of Object.entries(expected.cells)) {
    const got = (await cell(`${name}!${address}`)) ?? 0; // zero flows may be written as 0
    if (!close(got, want)) wrong.push(`${address}=${got}, wanted ${want}`);
  }
  check(`${name}: cells match the independent answer`, wrong.length === 0, wrong.join("; "));
  const body = await page.locator("body").innerText();
  check(`${name}: objective matches`, body.includes(String(+expected.objective.toFixed(4)).slice(0, 5)), body.slice(0, 300));
}
await page.screenshot({ path: join(SHOTS, "example-solved.png"), fullPage: true });

// Inserting twice gets a new sheet, not an overwrite.
await page.getByRole("tab", { name: "Model" }).click();
await page.getByRole("button", { name: "File" }).click();
await page.getByRole("menuitem", { name: "View Examples" }).click();
await page.getByRole("button", { name: "Insert Diet" }).click();
await page.getByText("Model for Diet 2").waitFor();
check("a second insert uses a new sheet", true);
await page.waitForTimeout(500);
const savedExample = await page.evaluate(() => JSON.parse(localStorage.getItem("gs-mock-state")));
check("an inserted example is saved", Object.values(savedExample.sheets).some((s) => s.name === "Diet 2" && s.model?.includes('Model("diet")')));
await page.evaluate(() => {
  localStorage.removeItem("gs-mock-state");
});
await openPane();

// ---- SolverStudio-style model (acceptance 6): in-place changes and reassignment both write back
await setCells({ A1: "p", A2: "q", B1: 2, B2: 3, E1: 4 });
await setItems([
  item("price", "indexed", "B1:B2", ["A1:A2"]),
  item("result", "indexed", "C1:C2", ["A1:A2"]),
  item("factor", "single", "E1"),
  item("names", "list", "A1:A2", [], true),
]);
await setModel(`from SolverStudio import *
import SolverStudio
from gurobipy import *

m = Model()
y = m.addVars(price.keys(), obj=price)
m.addConstrs(y[k] >= 1 for k in price)
m.optimize()
for k in price:
    result[k] = y[k].X * factor    # changed in place
factor = factor + 1                # reassigned
names.append("r")                  # read-only: warns, isn't written
print("shim", SolverStudio.price is price)
`);
const style = await solveAndApprove();
check("SolverStudio-style model solves", style === "Optimal", style);
check("in-place changes write back", (await cell("C1")) === 4 && (await cell("C2")) === 4);
check("reassignment writes back", (await cell("E1")) === 5);
check("read-only item isn't written", (await cell("A1")) === "p" && (await page.getByText("names changed but is read-only").first().isVisible()));
check("the shim sees the same objects", (await log()).includes("shim True"));
await setItems([]);

// ---- Dark theme follows the system
await page.emulateMedia({ colorScheme: "dark" });
await page.getByRole("tab", { name: "Model" }).click();
await page.waitForTimeout(200);
const darkBg = await page.locator(".fui-FluentProvider").first().evaluate((e) => getComputedStyle(e).backgroundColor);
const rgb = darkBg.match(/\d+/g).slice(0, 3).map(Number);
check("dark theme applies", Math.max(...rgb) < 80, darkBg);
await page.screenshot({ path: join(SHOTS, "dark.png"), fullPage: true });
await page.emulateMedia({ colorScheme: "light" });

//__NEXT__
check("no uncaught page errors", pageErrors.length === 0, pageErrors.join(" | "));
await browser.close();
helper.kill();
rmSync(data, { recursive: true, force: true });
console.log(failures ? `\n${failures} FAILED` : "\nAll passed");
console.log(`Screenshots in ${SHOTS}`);
process.exit(failures ? 1 : 0);
