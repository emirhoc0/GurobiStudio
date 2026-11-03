// Big data, end to end: real pane (mock sheet) -> real helper -> real Gurobi. Checks answers, not just
// that something finished. The sheet is the mock, so Excel's own speed is measured separately in
// docs/TESTING.md. Needs `npm run build`. Set GS_BIG_SCALE=0.1 for a quick run.
import { execSync, spawn, spawnSync } from "node:child_process";
import { mkdtempSync, readFileSync, readdirSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { webkit } from "playwright";

const PORT = 8798;
const ROOT = new URL("../../", import.meta.url).pathname;
const PYTHON = join(ROOT, "helper/.venv/bin/python");
const SCALE = Number(process.env.GS_BIG_SCALE ?? 1);
const scaled = (n) => Math.max(2, Math.round(n * Math.sqrt(SCALE)));
const data = mkdtempSync(join(tmpdir(), "gs-big-"));
const helper = spawn(PYTHON, ["-m", "gurobistudio"], {
  cwd: join(ROOT, "helper"),
  env: { ...process.env, GS_PORT: String(PORT), GS_DATA_DIR: data },
  stdio: ["ignore", "ignore", "inherit"],
});
process.on("exit", () => helper.kill());

let failures = 0;
const check = (name, ok, detail = "") => {
  console.log(`${ok ? "PASS" : "FAIL"}  ${name}${ok ? "" : `  -> ${detail}`}`);
  if (!ok) failures++;
};
const near = (a, b) => Math.abs(a - b) <= 1e-5 * Math.max(1, Math.abs(b));

process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";
for (let i = 0; ; i++) {
  try {
    await fetch(`https://localhost:${PORT}/`);
    break;
  } catch (e) {
    if (i > 100) throw e;
    await new Promise((r) => setTimeout(r, 100));
  }
}
const key = readFileSync(join(data, "page_key"), "utf8").trim();
const browser = await webkit.launch();
const page = await (await browser.newContext({ ignoreHTTPSErrors: true, viewport: { width: 380, height: 900 } })).newPage();
page.on("pageerror", (e) => check("no page errors", false, e.message));
await page.goto(`https://localhost:${PORT}/?k=${key}`);
await page.getByRole("tab", { name: "Model" }).waitFor();
await page.evaluate(() => {
  // Times the adapter's two big calls, and adds the column helper the fill functions need.
  window.timing = {};
  const adapter = window.mockExcel.adapter;
  for (const name of ["readRefs", "writeRanges"]) {
    const original = adapter[name];
    adapter[name] = async (...args) => {
      const t0 = performance.now();
      const out = await original(...args);
      window.timing[name] = (performance.now() - t0) / 1000;
      return out;
    };
  }
  window.col = (n) => (n === 0 ? "" : window.col(Math.floor((n - 1) / 26)) + String.fromCharCode(65 + ((n - 1) % 26)));
});

let peakMb = 0;
const sampler = setInterval(() => {
  try {
    const pids = [helper.pid, ...execSync(`pgrep -P ${helper.pid} || true`).toString().split("\n").filter(Boolean)];
    const kb = execSync(`ps -o rss= -p ${pids.join(",")}`).toString().trim().split(/\s+/).map(Number);
    peakMb = Math.max(peakMb, kb.reduce((a, b) => a + b, 0) / 1024);
  } catch {
    // a process ended between the two commands
  }
}, 300);

const item = (name, type, cellRange, indexRanges = [], readOnly = false) => ({ name, type, cellRange, indexRanges, readOnly });
const setUp = (model, items) =>
  page.evaluate(
    ([m, i]) => {
      window.gsModels.set({ ...window.gsModels.get(), "mock-sheet": m });
      window.gsItems.set({ ...window.gsItems.get(), "mock-sheet": i });
    },
    [model, items],
  );
const badge = () => page.locator(".fui-Badge").last().innerText();
const running = () => page.evaluate(() => /\bRunning\b/.test(document.body.innerText));
async function waitDone() {
  await page.waitForFunction(() => !/\bRunning\b/.test(document.body.innerText) && document.querySelector('[data-testid="log"]'), null, { timeout: 900_000 });
}
const clickSolve = async () => {
  await page.getByRole("tab", { name: "Output" }).click();
  await page.getByRole("button", { name: "Solve Model" }).click();
};
async function solve(firstTime = true) {
  const t0 = Date.now();
  await clickSolve();
  if (firstTime) await page.getByRole("alertdialog").getByRole("button", { name: "Run" }).click({ timeout: 300_000 });
  await waitDone();
  const seconds = (Date.now() - t0) / 1000;
  const text = await page.evaluate(() => document.body.innerText);
  const timing = await page.evaluate(() => window.timing);
  return { seconds, status: await badge(), text, timing, objective: Number(/Objective (-?[\d.e+-]+)\. /.exec(text)?.[1]), elapsed: Number(/([\d.]+)s\. /.exec(text)?.[1]) };
}
const rows = [];
const report = (name, cells, written, r) => {
  rows.push(`${name.padEnd(12)} ${String(cells).padStart(9)} ${String(written).padStart(8)} ${(r.timing.readRefs ?? 0).toFixed(1).padStart(7)} ${String(r.elapsed).padStart(8)} ${(r.timing.writeRanges ?? 0).toFixed(1).padStart(7)} ${r.seconds.toFixed(1).padStart(8)} ${peakMb.toFixed(0).padStart(8)}`);
  peakMb = 0;
};
const runDirs = () => readdirSync(join(data, "runs")).length;

// ---- Production plan: the shape of a real workbook, scaled up. Input-heavy.
const P = scaled(2000);
const R = scaled(500);
const lastCol = await page.evaluate((p) => window.col(p + 1), P);
const capCol = await page.evaluate((p) => window.col(p + 2), P);
await page.evaluate(
  ([P, R]) => {
    const cells = window.mockCells;
    cells.clear();
    for (let c = 1; c <= P; c++) {
      const column = window.col(c + 1);
      cells.set(`${column}1`, `P${c}`);
      for (let r = 1; r <= R; r++) cells.set(`${column}${r + 1}`, 1 + ((r * 31 + c * 17) % 97) / 10);
      cells.set(`${column}${R + 2}`, 5 + ((c * 13) % 41));
    }
    const capacity = window.col(P + 2);
    for (let r = 1; r <= R; r++) {
      cells.set(`A${r + 1}`, `R${r}`);
      cells.set(`${capacity}${r + 1}`, 10000 + ((r * 7) % 101) * 100);
    }
  },
  [P, R],
);
const planItems = [
  item("Sure", "table", `B2:${lastCol}${R + 1}`, [`A2:A${R + 1}`, `B1:${lastCol}1`], true),
  item("Kar", "indexed", `B${R + 2}:${lastCol}${R + 2}`, [`B1:${lastCol}1`], true),
  item("Kapasite", "indexed", `${capCol}2:${capCol}${R + 1}`, [`A2:A${R + 1}`], true),
  item("Uretim", "indexed", `B${R + 3}:${lastCol}${R + 3}`, [`B1:${lastCol}1`]),
];
const planModel = `from SolverStudio import *
from gurobipy import *
urunler = list(Kar)
m = Model()
x = m.addVars(urunler)
m.setObjective(quicksum(Kar[p] * x[p] for p in urunler), GRB.MAXIMIZE)
for r in Kapasite:
    m.addConstr(quicksum(Sure[r, p] * x[p] for p in urunler) <= Kapasite[r])
m.optimize()
Uretim = {p: x[p].X for p in urunler}
`;
const inputsChecksum = () =>
  page.evaluate(
    ([P, R]) => {
      let sum = 0;
      for (let c = 1; c <= P; c++) for (let r = 1; r <= R + 1; r++) sum += window.mockCells.get(`${window.col(c + 1)}${r + 1}`);
      return sum;
    },
    [P, R],
  );
const before = await inputsChecksum();
await setUp(planModel, planItems);
const plan = await solve();
check(`plan ${P}x${R}: Optimal`, plan.status === "Optimal", `${plan.status}: ${plan.text.slice(0, 600)}`);
check(`plan: Gurobi saw ${P * R} nonzeros`, plan.text.includes(`${P * R} nonzeros`), plan.text.slice(0, 600));

const direct = spawnSync(
  PYTHON,
  ["-c", `
from gurobipy import *
P, R = ${P}, ${R}
m = Model()
m.Params.OutputFlag = 0
x = m.addVars(range(1, P + 1))
m.setObjective(quicksum((5 + c * 13 % 41) * x[c] for c in x), GRB.MAXIMIZE)
for r in range(1, R + 1):
    m.addConstr(quicksum((1 + ((r * 31 + c * 17) % 97) / 10) * x[c] for c in x) <= 10000 + (r * 7 % 101) * 100)
m.optimize()
print("OBJ", m.ObjVal)
`],
  { encoding: "utf8" },
);
const directObjective = Number(/OBJ (\S+)/.exec(direct.stdout)?.[1]);
check("plan: objective matches gurobipy run directly", near(plan.objective, directObjective), `${plan.objective} vs ${directObjective} ${direct.stderr.slice(-300)}`);
const profit = await page.evaluate(
  ([P, R]) => {
    let sum = 0;
    for (let c = 1; c <= P; c++) sum += (window.mockCells.get(`${window.col(c + 1)}${R + 2}`) ?? 0) * (window.mockCells.get(`${window.col(c + 1)}${R + 3}`) ?? 0);
    return sum;
  },
  [P, R],
);
check("plan: cells written to the sheet reproduce the objective", near(profit, plan.objective), `${profit} vs ${plan.objective}`);
check("plan: read-only input cells unchanged", (await inputsChecksum()) === before);
report("plan", P * R + P + R, P, plan);

const again = await solve(false);
check("plan: a second solve finds nothing to change", /No items changed/.test(again.text), again.text.slice(0, 300));

// ---- Stop, at each stage, on the big plan.
const stopButton = () => page.getByRole("button", { name: "Stop" });
const dirs = runDirs();
await clickSolve();
await stopButton().click();
await waitDone();
check("stop during the read: Stopped", (await badge()) === "Stopped", await badge());
console.log(`      (${runDirs() === dirs ? "stopped before a run began" : "the read had finished, so the run was cancelled"})`);

const stopOnLog = async (label, model, marker) => {
  await setUp(model, planItems);
  await clickSolve();
  await page.getByRole("alertdialog").getByRole("button", { name: "Run" }).click({ timeout: 300_000 });
  await page.waitForFunction((m) => document.querySelector('[data-testid="log"]')?.innerText.includes(m), marker, { timeout: 300_000 });
  const t0 = Date.now();
  if (!(await running())) {
    check(`stop during ${label}`, false, `already ended: ${await badge()}`);
    return;
  }
  await page.getByRole("button", { name: "Stop" }).click();
  await waitDone();
  const seconds = (Date.now() - t0) / 1000;
  check(`stop during ${label}: Stopped within 6 s (${seconds.toFixed(1)} s)`, (await badge()) === "Stopped" && seconds < 6, `${await badge()}, ${seconds}`);
};
await stopOnLog("model code", `import time\nprint("kuruluyor", flush=True)\ntime.sleep(600)\n`, "kuruluyor");
await stopOnLog(
  "a Gurobi solve",
  planModel
    .replace("m.addVars(urunler)", "m.addVars(urunler, vtype=GRB.BINARY)") // a tight knapsack stays hard
    .replace("<= Kapasite[r]", "<= 0.25 * sum(Sure[r, p] for p in urunler)"),
  "Optimize a model",
);

// ---- Transportation: output-heavy, every cell written.
const S = scaled(400);
const C = scaled(250);
const flowLast = await page.evaluate((c) => window.col(c + 1), C);
await page.evaluate(
  ([S, C]) => {
    const cells = window.mockCells;
    cells.clear();
    for (let c = 1; c <= C; c++) {
      const column = window.col(c + 1);
      cells.set(`${column}1`, `C${c}`);
      for (let s = 1; s <= S; s++) cells.set(`${column}${s + 1}`, 1 + ((s * 7 + c * 13) % 50));
      cells.set(`${column}${S + 2}`, Math.floor(1600 / C * 10 * (S / 400)) + (c % 20));
    }
    for (let s = 1; s <= S; s++) {
      cells.set(`A${s + 1}`, `S${s}`);
      cells.set(`A${s + S + 3}`, `S${s}`);
      cells.set(`${window.col(C + 2)}${s + 1}`, 60);
    }
  },
  [S, C],
);
const supplyCol = await page.evaluate((c) => window.col(c + 2), C);
const flowTop = S + 4;
const flowBottom = S + S + 3;
await setUp(
  `from SolverStudio import *
from gurobipy import *
m = Model()
x = m.addVars(Cost.keys(), obj=Cost)
m.addConstrs((x.sum(s, "*") <= Supply[s] for s in Supply))
m.addConstrs((x.sum("*", c) >= Demand[c] for c in Demand))
m.optimize()
Flow = {k: v.X for k, v in x.items()}
`,
  [
    item("Cost", "table", `B2:${flowLast}${S + 1}`, [`A2:A${S + 1}`, `B1:${flowLast}1`], true),
    item("Demand", "indexed", `B${S + 2}:${flowLast}${S + 2}`, [`B1:${flowLast}1`], true),
    item("Supply", "indexed", `${supplyCol}2:${supplyCol}${S + 1}`, [`A2:A${S + 1}`], true),
    item("Flow", "table", `B${flowTop}:${flowLast}${flowBottom}`, [`A${flowTop}:A${flowBottom}`, `B1:${flowLast}1`]),
  ],
);
const flow = await solve();
check(`transport ${S}x${C}: Optimal`, flow.status === "Optimal", `${flow.status}: ${flow.text.slice(0, 600)}`);
const flowCheck = await page.evaluate(
  ([S, C, top]) => {
    let cost = 0;
    let written = 0;
    for (let c = 1; c <= C; c++) {
      const column = window.col(c + 1);
      for (let s = 1; s <= S; s++) {
        const v = window.mockCells.get(`${column}${top + s - 1}`);
        if (v !== undefined) written++;
        cost += window.mockCells.get(`${column}${s + 1}`) * (v ?? 0);
      }
    }
    return { cost, written };
  },
  [S, C, flowTop],
);
check("transport: every result cell written", flowCheck.written === S * C, `${flowCheck.written} of ${S * C}`);
check("transport: written cells reproduce the objective", near(flowCheck.cost, flow.objective), `${flowCheck.cost} vs ${flow.objective}`);
report("transport", S * C + S + C, S * C, flow);

// ---- A model that prints a line per row must not freeze the pane.
const LINES = Math.round(1_000_000 * SCALE);
await setUp(`for i in range(${LINES}):\n    print("satir", i)\nout = 1\n`, [item("out", "single", "A1")]);
await page.evaluate(() => window.mockCells.clear());
await clickSolve();
await page.getByRole("alertdialog").getByRole("button", { name: "Run" }).click();
let worst = 0;
while (await running()) {
  const t0 = Date.now();
  await page.evaluate(() => 1);
  worst = Math.max(worst, Date.now() - t0);
  await new Promise((r) => setTimeout(r, 100));
}
await waitDone();
const logText = await page.getByTestId("log").innerText();
check(`chatty model: pane answered within 1 s throughout (worst ${worst} ms)`, worst < 1000, `${worst} ms`);
check("chatty model: finished and the result arrived", (await page.evaluate(() => window.mockCells.get("A1"))) === 1);
check("chatty model: the log shows the end of the output", logText.includes(`satir ${LINES - 1}`), logText.slice(-200));
const newest = readdirSync(join(data, "runs")).sort((a, b) => Number(a.split("-")[0]) - Number(b.split("-")[0])).at(-1);
const onDisk = readFileSync(join(data, "runs", newest, "log.txt"), "utf8").split("\n").length;
check("chatty model: every line is in log.txt", onDisk >= LINES, `${onDisk}`);
rows.push(`${"chatty".padEnd(12)} ${LINES} log lines, peak helper+runner ${peakMb.toFixed(0)} MB`);

clearInterval(sampler);
console.log(`\n${"scenario".padEnd(12)} ${"cells in".padStart(9)} ${"written".padStart(8)} ${"read s".padStart(7)} ${"solve s".padStart(8)} ${"write s".padStart(7)} ${"total s".padStart(8)} ${"peak MB".padStart(8)}`);
rows.forEach((r) => console.log(r));
console.log(failures ? `\n${failures} failed` : "\nAll passed");
await browser.close();
process.exit(failures ? 1 : 0);
