import type { Cell, ExcelAdapter, RangeRef, RangeWrite } from "./adapter";

function columnNumber(letters: string): number {
  return [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
}

function columnLetters(n: number): string {
  return n === 0
    ? ""
    : columnLetters(Math.floor((n - 1) / 26)) + String.fromCharCode(65 + ((n - 1) % 26));
}

function cellsOf(address: string): string[][] {
  const match = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(address);
  if (!match) throw new Error(`Bad address: ${address}`);
  const [, c1, r1, c2 = c1, r2 = r1] = match;
  const rows: string[][] = [];
  for (let r = Number(r1); r <= Number(r2); r++) {
    const row: string[] = [];
    for (let c = columnNumber(c1); c <= columnNumber(c2); c++) row.push(columnLetters(c) + r);
    rows.push(row);
  }
  return rows;
}

/** In-memory workbook, so the UI runs in a normal browser. */
export function createMockAdapter(initial: Record<string, Cell> = {}) {
  const cells = new Map(Object.entries(initial));
  const adapter: ExcelAdapter = {
    activeSheet: async () => ({ id: "mock-sheet", name: "Sheet1" }),
    readRanges: async (refs: RangeRef[]) =>
      refs.map((r) => cellsOf(r.address).map((row) => row.map((a) => cells.get(a) ?? ""))),
    writeRanges: async (writes: RangeWrite[]) => {
      for (const w of writes) {
        cellsOf(w.ref.address).forEach((row, i) =>
          row.forEach((a, j) => {
            const value = w.values[i][j];
            if (value === null) return;
            if (value === "") cells.delete(a);
            else cells.set(a, value);
          }),
        );
      }
    },
  };
  return { adapter, cells };
}
