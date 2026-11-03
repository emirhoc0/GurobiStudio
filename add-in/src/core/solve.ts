// ponytail: M2 proof with one hard-coded model and cells; M3 and M4 replace them with real items.
import type { ExcelAdapter } from "../excel/adapter";
import { runStore, type RunError } from "../ui/store";
import { apiFetch } from "./api";
import { createSseParser } from "./sse";

const MODEL = `from SolverStudio import *
from gurobipy import *

m = Model("proof")
x = m.addVar(ub=capacity, name="x")
m.setObjective(2 * x, GRB.MAXIMIZE)
m.optimize()
result = x.x
print("x =", result)
`;

type RunEvent =
  | { type: "log"; line: string }
  | { type: "status"; state: "running" | "stopped" }
  | {
      type: "result";
      elapsed: number;
      status: string;
      objective: number | null;
      changed: { name: string; value: unknown }[];
      warnings: string[];
    }
  | { type: "error"; elapsed: number; error: RunError };

function fail(message: string, type = "Error") {
  runStore.set({ ...runStore.get(), phase: "failed", error: { type, message, line: null } });
}

export async function solve(excel: ExcelAdapter) {
  if (runStore.get().phase === "running") return;
  runStore.set({ phase: "running", log: [] });
  try {
    const sheet = await excel.activeSheet();
    const capacityCell = { sheetId: sheet.id, address: "A1" };
    const resultCell = { sheetId: sheet.id, address: "B1" };

    const [[[stored]]] = await excel.readRanges([capacityCell]);
    if (stored === "") await excel.writeRanges([{ ref: capacityCell, values: [[10]] }]);
    const capacity = stored === "" ? 10 : stored;
    if (typeof capacity !== "number") {
      fail("Cell A1 must hold a number for this test model.");
      return;
    }

    const start = await apiFetch("/api/run", {
      method: "POST",
      body: JSON.stringify({
        model: MODEL,
        items: [
          { name: "capacity", type: "single", value: capacity },
          { name: "result", type: "single", value: null },
        ],
        params: {},
        workbook_id: "m2-proof",
        approved: true, // the trust prompt arrives in M4
      }),
    });
    if (start.status === 409) {
      fail("Another solve is still running.");
      return;
    }
    if (!start.ok) throw new Error(`Helper returned ${start.status}`);
    const { run_id: runId } = await start.json();
    runStore.set({ ...runStore.get(), runId });

    const stream = await apiFetch(`/api/run/${runId}/events`);
    if (!stream.ok || !stream.body) throw new Error(`Helper returned ${stream.status}`);
    const events: RunEvent[] = [];
    const feed = createSseParser<RunEvent>((event) => {
      if (event.type === "log") {
        const current = runStore.get();
        runStore.set({ ...current, log: [...current.log, event.line] });
      } else {
        events.push(event);
      }
    });
    const reader = stream.body.getReader();
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      feed(value);
    }

    const final = events[events.length - 1];
    const current = runStore.get();
    if (final?.type === "status" && final.state === "stopped") {
      runStore.set({ ...current, phase: "stopped" });
    } else if (final?.type === "error") {
      runStore.set({ ...current, phase: "failed", elapsed: final.elapsed, error: final.error });
    } else if (final?.type === "result") {
      const written = final.changed.find((c) => c.name === "result");
      if (written) {
        await excel.writeRanges([{ ref: resultCell, values: [[written.value as number]] }]);
      }
      runStore.set({
        ...current,
        phase: "finished",
        status: final.status,
        objective: final.objective,
        elapsed: final.elapsed,
        changed: final.changed.map((c) => c.name),
        warnings: final.warnings,
      });
    } else {
      throw new Error("The helper closed the stream without a result.");
    }
  } catch (e) {
    fail(`Could not reach the helper. Is it running? (${(e as Error).message})`);
  }
}

export async function stop() {
  const { phase, runId } = runStore.get();
  if (phase === "running" && runId) await apiFetch(`/api/run/${runId}/cancel`, { method: "POST" });
}
