// Reads the active sheet's data items, runs its model, writes changed items back.
import type { ExcelAdapter, RangeWrite } from "../excel/adapter";
import { itemsStore, MODEL_TEMPLATE, modelStore, runStore, type RunError } from "../ui/store";
import { apiFetch } from "./api";
import { overlapWarnings, readItems, type BuiltItem, type Wire } from "./shapes";
import { createSseParser } from "./sse";
import { writeFor } from "./writeback";

type RunEvent =
  | { type: "log"; line: string }
  | { type: "status"; state: "running" | "stopped" }
  | {
      type: "result";
      elapsed: number;
      status: string;
      objective: number | null;
      changed: { name: string; value: Wire }[];
      warnings: string[];
    }
  | { type: "error"; elapsed: number; error: RunError };

function fail(message: string, type = "Error") {
  runStore.set({ ...runStore.get(), phase: "failed", error: { type, message, line: null } });
}

/** Null after reporting why the items can't be read. */
async function readForRun(excel: ExcelAdapter, sheetId: string): Promise<BuiltItem[] | null> {
  const results = await readItems(excel, itemsStore.get()[sheetId] ?? [], sheetId);
  const problems = results.flatMap((r) => ("error" in r ? [r.error] : []));
  if (problems.length) {
    fail(problems.join("\n"), "Data item problem");
    return null;
  }
  return results as BuiltItem[];
}

function blankNotes(built: BuiltItem[]): string[] {
  return built
    .filter((b) => b.blanks > 0 && b.slots)
    .map((b) => `${b.item.name}: ${b.blanks} blank cell(s) omitted`);
}

export async function solve(excel: ExcelAdapter) {
  if (runStore.get().phase === "running") return;
  runStore.set({ phase: "running", log: [] });

  let built: BuiltItem[];
  let sheetId: string;
  try {
    sheetId = (await excel.activeSheet()).id;
    const read = await readForRun(excel, sheetId);
    if (!read) return;
    built = read;
  } catch (e) {
    fail(`Could not read the sheet. (${(e as Error).message})`);
    return;
  }
  const readWarnings = overlapWarnings(built);
  runStore.set({ ...runStore.get(), log: blankNotes(built), warnings: readWarnings });

  try {
    const start = await apiFetch("/api/run", {
      method: "POST",
      body: JSON.stringify({
        model: modelStore.get()[sheetId] ?? MODEL_TEMPLATE,
        items: built.map((b) => ({
          name: b.item.name,
          type: b.item.type,
          value: b.value,
          readOnly: b.item.readOnly,
        })),
        params: {},
        workbook_id: "m3-unsaved",
        approved: true, // the model is typed here until M5; the trust prompt arrives in M4
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
      const writes: RangeWrite[] = [];
      const warnings = [...readWarnings, ...final.warnings];
      for (const change of final.changed) {
        const item = built.find((b) => b.item.name === change.name);
        if (!item) continue;
        const result = writeFor(item, change.value);
        if (result.write) writes.push(result.write);
        warnings.push(...result.warnings);
      }
      try {
        if (writes.length) await excel.writeRanges(writes);
      } catch (e) {
        fail(`The solve finished, but its results could not be written. (${(e as Error).message})`);
        return;
      }
      runStore.set({
        ...current,
        phase: "finished",
        status: final.status,
        objective: final.objective,
        elapsed: final.elapsed,
        changed: final.changed.map((c) => c.name),
        warnings,
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
