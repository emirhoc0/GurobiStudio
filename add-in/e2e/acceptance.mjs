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

//__NEXT__
check("no uncaught page errors", pageErrors.length === 0, pageErrors.join(" | "));
await browser.close();
helper.kill();
rmSync(data, { recursive: true, force: true });
console.log(failures ? `\n${failures} FAILED` : "\nAll passed");
console.log(`Screenshots in ${SHOTS}`);
process.exit(failures ? 1 : 0);
