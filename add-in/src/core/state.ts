// What's saved in the workbook. Pure, so it can be tested without Excel.
import type { DataItem } from "./shapes";

export const SCHEMA = 1;

export interface Settings {
  timeLimit: string;
  mipGap: string;
  threads: string;
  outputFlag: boolean;
  custom: { key: string; value: string }[];
}

export interface SheetState {
  name: string;
  language: "gurobi"; // the only hook kept for adding PuLP later
  model?: string;
  items: DataItem[];
}

export interface WorkbookState {
  schema: number;
  workbookId: string;
  settings: Settings;
  sheets: Record<string, SheetState>;
}

export const DEFAULT_SETTINGS: Settings = {
  timeLimit: "",
  mipGap: "",
  threads: "",
  outputFlag: true,
  custom: [],
};

/** Throws when the text isn't state this version understands, so it's never overwritten. */
export function parseState(text: string): WorkbookState {
  const state = JSON.parse(text) as WorkbookState;
  if (typeof state?.schema !== "number" || typeof state.workbookId !== "string") {
    throw new Error("It is not GurobiStudio state.");
  }
  if (state.schema > SCHEMA) {
    throw new Error(`It was saved by a newer version (schema ${state.schema}).`);
  }
  return {
    ...state,
    settings: { ...DEFAULT_SETTINGS, ...state.settings },
    sheets: state.sheets ?? {},
  };
}

/**
 * Worksheet ids should survive a reopen, but if none of them match, the file lost them, so match by
 * name. If any id matches, ids are fine and an unmatched entry belongs to a deleted sheet, which a
 * new sheet with the same name must not inherit.
 */
export function remapSheets(
  sheets: Record<string, SheetState>,
  current: { id: string; name: string }[],
): Record<string, SheetState> {
  const ids = Object.keys(sheets);
  if (!ids.length || ids.some((id) => current.some((s) => s.id === id))) return sheets;
  return Object.fromEntries(
    ids.map((id) => [current.find((s) => s.name === sheets[id].name)?.id ?? id, sheets[id]]),
  );
}

/** Blank fields are left to Gurobi. The table comes last, so it can override the fields. */
export function buildParams(settings: Settings): Record<string, string | number> {
  const params: Record<string, string | number> = {};
  const put = (key: string, text: string) => {
    const value = text.trim();
    if (!key || !value) return;
    params[key] = Number.isNaN(Number(value)) ? value : Number(value);
  };
  put("TimeLimit", settings.timeLimit);
  put("MIPGap", settings.mipGap);
  put("Threads", settings.threads);
  if (!settings.outputFlag) params.OutputFlag = 0;
  for (const { key, value } of settings.custom) put(key.trim(), value);
  return params;
}
