import { columnLetters, freeSheetName, parseRef, rectOf } from "../core/refs";
import { MAX_RANGE_CELLS, tooManyCells } from "./adapter";
import type { Cell, ExcelAdapter, RangeRef, RangeWrite, RefResult } from "./adapter";

const ERROR_TEXT = /^#(N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!)$/;
const MOCK_STATE_KEY = "gs-mock-state";
const storage = () => (typeof localStorage === "undefined" ? null : localStorage); // reload = reopen

function cellsOf(address: string): string[][] {
  const { r1, c1, r2, c2 } = rectOf(address);
  const rows: string[][] = [];
  for (let r = r1; r <= r2; r++) {
    const row: string[] = [];
    for (let c = c1; c <= c2; c++) row.push(columnLetters(c) + r);
    rows.push(row);
  }
  return rows;
}

export interface MockTable {
  /** Data body, without the header row. */
  address: string;
  columns: string[];
}

export interface MockOptions {
  /** A name maps to an address, or to values when it is a formula name with no range. */
  names?: Record<string, string | Cell[][]>;
  tables?: Record<string, MockTable>;
}

/** In-memory workbook, so the UI runs in a normal browser. Cells on Sheet2 use "Sheet2!A1" keys. */
export function createMockAdapter(initial: Record<string, Cell> = {}, options: MockOptions = {}) {
  const cells = new Map(Object.entries(initial));
  const names = new Map(Object.entries(options.names ?? {}));
  const tables = new Map(Object.entries(options.tables ?? {}));
  const formulas = new Map<string, string>(); // tests add a cell's formula here
  const boldRanges: string[] = [];
  const listeners: (() => void)[] = [];
  const sheets = [
    { id: "mock-sheet", name: "Sheet1" },
    { id: "mock-sheet-2", name: "Sheet2" },
  ];
  let active = sheets[0];
  let selected = "Sheet1!A1";

  const sheetOf = (name: string | undefined, sheetId: string) =>
    name ? sheets.find((s) => s.name === name)?.id : sheetId;
  const key = (sheetId: string, address: string) => {
    const other = sheets.slice(1).find((s) => s.id === sheetId);
    return other ? `${other.name}!${address}` : address;
  };

  const read = (ref: RangeRef): RefResult => {
    const { r1, c1, r2, c2 } = rectOf(ref.address);
    const size = (r2 - r1 + 1) * (c2 - c1 + 1);
    if (size > MAX_RANGE_CELLS) return { error: tooManyCells(ref.address, size) };
    const values = cellsOf(ref.address).map((row) =>
      row.map((a) => cells.get(key(ref.sheetId, a)) ?? ""),
    );
    return { ref, values, errorCells: errorCells(values) };
  };

  const errorCells = (values: Cell[][]) =>
    values.flatMap((row, r) =>
      row.flatMap((v, c): [number, number][] =>
        typeof v === "string" && ERROR_TEXT.test(v) ? [[r, c]] : [],
      ),
    );

  const readOne = (text: string, sheetId: string): RefResult => {
    try {
      const p = parseRef(text);
      if (p.kind === "a1") {
        const id = sheetOf(p.sheet, sheetId);
        if (!id) return { error: `Could not read ${text}: no such sheet.` };
        return read({ sheetId: id, address: p.address });
      }
      if (p.kind === "table") {
        const table = tables.get(p.table);
        const [from, to] = [p.first, p.last].map((c) => table?.columns.indexOf(c) ?? -1);
        if (!table || from < 0 || to < 0) return { error: `Could not read ${text}.` };
        const { r1, c1, r2 } = rectOf(table.address);
        const address = `${columnLetters(c1 + from)}${r1}:${columnLetters(c1 + to)}${r2}`;
        return read({ sheetId, address });
      }
      const named = names.get(p.name);
      if (Array.isArray(named)) return { ref: null, values: named, errorCells: errorCells(named) };
      if (named) return readOne(named, sheetId);
      const table = tables.get(p.name);
      if (table) return read({ sheetId, address: table.address });
      return { error: `There is no range, name or table called ${p.name}.` };
    } catch (e) {
      return { error: (e as Error).message };
    }
  };

  const adapter: ExcelAdapter = {
    activeSheet: async () => active,
    onSheetActivated: (listener) => {
      listeners.push(listener);
    },
    readRefs: async (refs, sheetId) => refs.map((r) => readOne(r, sheetId)),
    writeRanges: async (writes: RangeWrite[]) => {
      for (const w of writes) {
        cellsOf(w.ref.address).forEach((row, i) =>
          row.forEach((a, j) => {
            const value = w.values[i][j];
            if (value === null) return;
            formulas.delete(key(w.ref.sheetId, a));
            if (value === "") cells.delete(key(w.ref.sheetId, a));
            // Excel stores a leading apostrophe as "this is text", not as part of the value.
            else
              cells.set(
                key(w.ref.sheetId, a),
                String(value).startsWith("'") ? String(value).slice(1) : value,
              );
          }),
        );
      }
    },
    selection: async () => selected,
    listNames: async () => [...names.keys(), ...tables.keys()],
    createName: async (name, ref) => {
      if (names.has(name)) throw new Error(`The name ${name} already exists in Excel.`);
      if (/^[a-z]{1,3}\d+$/i.test(name) || /^[rc]$/i.test(name)) {
        throw new Error(`Excel does not allow ${name} as a name.`);
      }
      const sheet = sheets.find((s) => s.id === ref.sheetId)?.name ?? "Sheet1";
      names.set(name, `${sheet}!${ref.address}`);
    },
    formulaCells: async (writes) =>
      writes.map((w) =>
        cellsOf(w.ref.address).map((row, r) =>
          row.map((a, c) => w.values[r][c] !== null && formulas.has(key(w.ref.sheetId, a))),
        ),
      ),
    openEditorDialog: async (onMessage, onClosed) => {
      const popup = window.open(`${location.pathname}${location.search}&view=editor`, "_blank")!;
      const listen = (e: MessageEvent) => {
        if (e.source === popup) onMessage(e.data);
      };
      addEventListener("message", listen);
      const watch = setInterval(() => {
        if (!popup.closed) return;
        clearInterval(watch);
        removeEventListener("message", listen);
        onClosed();
      }, 200);
      return { send: (message) => popup.postMessage(message, location.origin) };
    },
    editorChannel: (onMessage) => {
      addEventListener("message", (e) => {
        if (e.source === window.opener) onMessage(e.data);
      });
      return (message) => window.opener.postMessage(message, location.origin);
    },
    loadState: async () => storage()?.getItem(MOCK_STATE_KEY) ?? null,
    saveState: async (text) => storage()?.setItem(MOCK_STATE_KEY, text),
    listSheets: async () => sheets,
    addSheet: async (name, rows, bold) => {
      const sheet = {
        id: `mock-sheet-${sheets.length + 1}`,
        name: freeSheetName(
          name,
          sheets.map((s) => s.name),
        ),
      };
      sheets.push(sheet);
      rows.forEach((row, r) =>
        row.forEach((value, c) => {
          if (value !== "") cells.set(key(sheet.id, columnLetters(c + 1) + (r + 1)), value);
        }),
      );
      bold.forEach((address) => boldRanges.push(`${sheet.name}!${address}`));
      return sheet;
    },
    activateSheet: async (sheetId) => {
      active = sheets.find((s) => s.id === sheetId) ?? active;
      listeners.forEach((l) => l());
    },
    officeDark: () => null,
  };

  return {
    adapter,
    cells,
    formulas,
    boldRanges,
    names,
    tables,
    select: (address: string) => {
      selected = address;
    },
    activate: (name: string) => {
      active = sheets.find((s) => s.name === name) ?? active;
      listeners.forEach((l) => l());
    },
  };
}
