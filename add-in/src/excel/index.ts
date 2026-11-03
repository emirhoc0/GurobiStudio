import type { ExcelAdapter } from "./adapter";

let current: ExcelAdapter | undefined;

export function setExcel(adapter: ExcelAdapter) {
  current = adapter;
}

export function excel(): ExcelAdapter {
  if (!current) throw new Error("The Excel adapter is not ready yet.");
  return current;
}
