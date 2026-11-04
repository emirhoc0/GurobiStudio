// Reads the active sheet's data items, runs its model, writes changed items back.
import type { ExcelAdapter, RangeWrite } from "../excel/adapter";
import {
  askConfirm,
  itemsStore,
  MODEL_TEMPLATE,
  modelStore,
  runStore,
  setTab,
  settingsStore,
  workbookIdStore,
  type RunError,
} from "../ui/store";
import { apiFetch, getHealth } from "./api";
import { whenLoaded } from "./persist";
import { overlapWarnings, readItems, type BuiltItem, type Wire } from "./shapes";
import { createSseParser } from "./sse";
import { buildParams } from "./state";
import { keepMasked, keepReadOnly, maskedAddresses, writeFor } from "./writeback";

// The full log is in the run folder; this keeps a chatty model from slowing the pane.
const MAX_LOG_LINES = 5000;
const LOG_FLUSH_MS = 200;

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

/** Overwrite confirmed cells or leave them. One prompt per solve, however many cells. */
async function protectFormulas(
  excel: ExcelAdapter,
  writes: RangeWrite[],
): Promise<{ writes: RangeWrite[]; warnings: string[] }> {
  if (!writes.length) return { writes, warnings: [] };
  const masks = await excel.formulaCells(writes);
  const cells = maskedAddresses(writes, masks);
  if (!cells.length) return { writes, warnings: [] };
  const shown = cells.slice(0, 8).join(", ") + (cells.length > 8 ? ", …" : "");
  const overwrite = await askConfirm({
    title: "Overwrite formulas?",
    body: `The results would replace ${cells.length} cell(s) that hold formulas: ${shown}. Keeping them still writes the other results.`,
    confirm: "Overwrite",
    cancel: "Keep formulas",
  });
  if (overwrite) return { writes, warnings: [] };
  return {
    writes: keepMasked(writes, masks),
    warnings: [`Kept ${cells.length} formula cell(s): ${shown}`],
  };
}

/** Text parameters can name a file or a server, and they come from the workbook, so show them. */
function withTextParams(model: string, params: Record<string, string | number>): string {
  const text = Object.entries(params).filter(([, value]) => typeof value === "string");
  if (!text.length) return model;
  const lines = text.map(([key, value]) => `# ${key} = ${JSON.stringify(value)}`);
  return `# Gurobi parameters set by this workbook:\n${lines.join("\n")}\n\n${model}`;
}

// Reading a large sheet takes a while, and Stop has no run to cancel until the helper names one.
let stopRequested = false;

const cancelRun = (runId: string) => apiFetch(`/api/run/${runId}/cancel`, { method: "POST" });
const markStopped = () => runStore.set({ ...runStore.get(), phase: "stopped" });

async function handleBusy(runId: string | null) {
  const stopIt =
    runId !== null &&
    (await askConfirm({
      title: "Another solve is running",
      body: "It may have started before this pane reloaded. Stop it? You can then solve again.",
      confirm: "Stop it",
      cancel: "Leave it running",
    }));
  if (!stopIt) {
    fail("Another solve is still running.");
    return;
  }
  await cancelRun(runId);
  markStopped();
}

export async function solve(excel: ExcelAdapter) {
  if (runStore.get().phase === "running") return;
  stopRequested = false;
  runStore.set({ phase: "running", log: [] });
  setTab("output");
  await whenLoaded();

  try {
    await getHealth();
  } catch (e) {
    fail(`Could not reach the helper. Is it running? (${(e as Error).message})`);
    return;
  }

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
  if (stopRequested) {
    markStopped();
    return;
  }
  const readWarnings = overlapWarnings(built);
  runStore.set({ ...runStore.get(), log: blankNotes(built), warnings: readWarnings });

  const model = modelStore.get()[sheetId] ?? MODEL_TEMPLATE;
  const params = buildParams(settingsStore.get());
  const startRun = (approved: boolean) =>
    apiFetch("/api/run", {
      method: "POST",
      body: JSON.stringify({
        model,
        items: built.map((b) => ({
          name: b.item.name,
          type: b.item.type,
          value: b.value,
          readOnly: b.item.readOnly,
        })),
        params,
        workbook_id: workbookIdStore.get(),
        approved,
      }),
    });

  try {
    let start = await startRun(false);
    if (start.status === 409 && (await start.clone().json()).trust_required) {
      const allowed = await askConfirm({
        title: "Run this model?",
        body: "Models run as Python with your full access. This code is new or has changed, so read it before running it.",
        code: withTextParams(model, params),
        confirm: "Run",
        cancel: "Cancel",
      });
      if (!allowed) {
        runStore.set({ phase: "idle", log: [] });
        return;
      }
      start = await startRun(true);
    }
    if (start.status === 409) {
      await handleBusy((await start.json()).run_id ?? null);
      return;
    }
    if (!start.ok) {
      const body = await start.json().catch(() => null);
      if (body?.error) {
        fail(body.error);
        return;
      }
      throw new Error(`Helper returned ${start.status}`);
    }
    const { run_id: runId } = await start.json();
    runStore.set({ ...runStore.get(), runId });
    if (stopRequested) await cancelRun(runId);

    const stream = await apiFetch(`/api/run/${runId}/events`);
    if (!stream.ok || !stream.body) throw new Error(`Helper returned ${stream.status}`);
    const events: RunEvent[] = [];
    let lines: string[] = [];
    const feed = createSseParser<RunEvent>((event) => {
      if (event.type === "log") lines.push(event.line);
      else events.push(event);
    });
    let flushTimer: ReturnType<typeof setTimeout> | undefined;
    const flushLog = () => {
      flushTimer = undefined;
      if (!lines.length) return;
      const shown = runStore.get();
      runStore.set({ ...shown, log: [...shown.log, ...lines].slice(-MAX_LOG_LINES) });
      lines = [];
    };
    const reader = stream.body.getReader();
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      feed(value);
      // On a timer, because rendering the log on every chunk freezes the pane on a chatty model.
      if (lines.length && flushTimer === undefined) flushTimer = setTimeout(flushLog, LOG_FLUSH_MS);
    }
    clearTimeout(flushTimer);
    flushLog();

    const final = events[events.length - 1];
    const current = runStore.get();
    if (final?.type === "status" && final.state === "stopped") {
      runStore.set({ ...current, phase: "stopped" });
    } else if (final?.type === "error") {
      runStore.set({ ...current, phase: "failed", elapsed: final.elapsed, error: final.error });
    } else if (final?.type === "result") {
      const proposed: RangeWrite[] = [];
      const warnings = [...readWarnings, ...final.warnings];
      for (const change of final.changed) {
        const item = built.find((b) => b.item.name === change.name);
        if (!item) continue;
        const result = writeFor(item, change.value);
        if (result.write) proposed.push(result.write);
        warnings.push(...result.warnings);
      }
      const guarded = keepReadOnly(proposed, built);
      warnings.push(...guarded.warnings);
      try {
        const kept = await protectFormulas(excel, guarded.writes);
        warnings.push(...kept.warnings);
        if (kept.writes.length) await excel.writeRanges(kept.writes);
      } catch (e) {
        fail(
          `The solve finished, but writing its results failed, and some of them may already be in the sheet. (${(e as Error).message})`,
        );
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
  if (phase !== "running") return;
  if (runId) await cancelRun(runId);
  else stopRequested = true;
}
