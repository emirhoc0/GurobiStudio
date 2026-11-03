// Turns cells into the JSON form of a Python value, and remembers which key lives in which cell.
import type { Cell, ExcelAdapter, RangeRef, RefRead, RefResult } from "../excel/adapter";
import { cellAt, overlaps, splitRefs } from "./refs";

export type ItemType = "single" | "list" | "indexed" | "table";

export interface DataItem {
  name: string;
  type: ItemType;
  cellRange: string;
  indexRanges: string[];
  readOnly: boolean;
}

export type Scalar = Cell | null;
export type Key = Cell | Cell[];
export type Wire = Scalar | (Scalar | Scalar[])[] | { __dict__: [Key, Scalar][] };

export interface BuiltItem {
  item: DataItem;
  /** Exactly what was read, so a table growing during a solve can't shift the results. */
  ref: RangeRef | null;
  old: Cell[][];
  value: Wire;
  count: number;
  blanks: number;
  /** Dict items only: every key in the index, including those whose cell is blank. */
  slots?: Map<string, [number, number]>;
}

export type ItemResult = BuiltItem | { item: DataItem; error: string };

export const keyId = (key: unknown) => JSON.stringify(key);
export const entries = (n: number) => `${n} ${n === 1 ? "entry" : "entries"}`;

const blankToNull = (v: Cell): Scalar => (v === "" ? null : v);

function lastNonBlank(lines: Cell[][]): number {
  for (let i = lines.length - 1; i >= 0; i--) {
    if (lines[i].some((v) => v !== "")) return i;
  }
  return -1;
}

const transpose = (m: Cell[][]): Cell[][] => m[0].map((_, c) => m.map((row) => row[c]));

function usable(item: DataItem, label: string, result: RefResult): RefRead {
  if ("error" in result) throw new Error(`${item.name}: ${result.error}`);
  if (result.errorCells.length) {
    const [r, c] = result.errorCells[0];
    const where = result.ref
      ? `cell ${cellAt(result.ref.address, r, c)}`
      : `row ${r + 1}, column ${c + 1}`;
    throw new Error(`${item.name}: ${label}${where} holds ${String(result.values[r][c])}.`);
  }
  return result;
}

/** Key components per position. Vertical means each row is one key. */
function indexKeys(item: DataItem, reads: RefRead[], vertical: boolean): Cell[][] {
  const lines = reads.map((r) => (vertical ? r.values : transpose(r.values)));
  if (new Set(lines.map((l) => l.length)).size > 1) {
    throw new Error(`${item.name}: the index ranges have different lengths.`);
  }
  const keys = lines[0].map((_, p) => lines.flatMap((l) => l[p]));
  const count = lastNonBlank(keys) + 1;
  if (count === 0) throw new Error(`${item.name}: the index range is empty.`);
  const used = keys.slice(0, count);
  used.forEach((key, p) => {
    if (key.includes("")) throw new Error(`${item.name}: index entry ${p + 1} has a blank cell.`);
  });
  return used;
}

function checkUnique(item: DataItem, keys: Key[]) {
  const seen = new Set<string>();
  for (const key of keys) {
    const id = keyId(key);
    if (seen.has(id)) throw new Error(`${item.name}: the index repeats ${id}.`);
    seen.add(id);
  }
}

/** The index decides how much is used: fewer data cells is an error, spare blank ones are fine. */
function checkFits(item: DataItem, lines: Cell[][], count: number) {
  if (lines.length < count) {
    throw new Error(
      `${item.name}: the index has ${entries(count)} but the cell range only fits ${lines.length}.`,
    );
  }
  if (lastNonBlank(lines) >= count) {
    throw new Error(
      `${item.name}: the cell range has values past the end of the index (${entries(count)}).`,
    );
  }
}

function buildList(read: RefRead): Pick<BuiltItem, "value" | "count" | "blanks"> {
  const rows = read.values.length;
  const cols = read.values[0].length;
  if (rows > 1 && cols > 1) {
    const used = read.values.slice(0, lastNonBlank(read.values) + 1);
    return { value: used.map((row) => row.map(blankToNull)), count: used.length, blanks: 0 };
  }
  const line = rows > 1 ? read.values.map((row) => [row[0]]) : read.values[0].map((v) => [v]);
  const used = line.slice(0, lastNonBlank(line) + 1).map(([v]) => blankToNull(v));
  return { value: used, count: used.length, blanks: 0 };
}

function buildIndexed(item: DataItem, cell: RefRead, indexes: RefRead[]) {
  const rows = cell.values.length;
  const cols = cell.values[0].length;
  if (rows > 1 && cols > 1) {
    throw new Error(`${item.name}: an Indexed item needs one row or column. Use Table for a grid.`);
  }
  const vertical = !(rows === 1 && cols > 1);
  for (const index of indexes) {
    const across = index.values.length === 1 && index.values[0].length > 1;
    const down = index.values.length > 1 && index.values[0].length === 1;
    if ((vertical && rows > 1 && across) || (!vertical && down)) {
      throw new Error(`${item.name}: the index and the cell range run in different directions.`);
    }
  }
  const keys = indexKeys(item, indexes, vertical).map((k): Key => (k.length === 1 ? k[0] : k));
  checkUnique(item, keys);
  const lines = vertical ? cell.values : transpose(cell.values);
  checkFits(item, lines, keys.length);

  const slots = new Map<string, [number, number]>();
  const dict: [Key, Scalar][] = [];
  keys.forEach((key, p) => {
    slots.set(keyId(key), vertical ? [p, 0] : [0, p]);
    if (lines[p][0] !== "") dict.push([key, lines[p][0]]);
  });
  return {
    value: { __dict__: dict },
    count: dict.length,
    blanks: keys.length - dict.length,
    slots,
  };
}

function vectorOf(item: DataItem, label: string, read: RefRead): Cell[] {
  const rows = read.values.length;
  const cols = read.values[0].length;
  if (rows > 1 && cols > 1) {
    throw new Error(`${item.name}: the ${label} index must be a single row or column.`);
  }
  const line = rows > 1 ? read.values.map((row) => row[0]) : read.values[0];
  const count = lastNonBlank(line.map((v) => [v])) + 1;
  if (count === 0) throw new Error(`${item.name}: the ${label} index is empty.`);
  const used = line.slice(0, count);
  const gap = used.indexOf("");
  if (gap >= 0) throw new Error(`${item.name}: ${label} index entry ${gap + 1} is blank.`);
  checkUnique(item, used);
  return used;
}

function buildTable(item: DataItem, cell: RefRead, rowIndex: RefRead, colIndex: RefRead) {
  const rowKeys = vectorOf(item, "row", rowIndex);
  const colKeys = vectorOf(item, "column", colIndex);
  checkFits(item, cell.values, rowKeys.length);
  checkFits(item, transpose(cell.values), colKeys.length);

  const slots = new Map<string, [number, number]>();
  const dict: [Key, Scalar][] = [];
  rowKeys.forEach((row, r) =>
    colKeys.forEach((col, c) => {
      slots.set(keyId([row, col]), [r, c]);
      if (cell.values[r][c] !== "") dict.push([[row, col], cell.values[r][c]]);
    }),
  );
  return {
    value: { __dict__: dict },
    count: dict.length,
    blanks: rowKeys.length * colKeys.length - dict.length,
    slots,
  };
}

/** Throws with a message that names the item. */
export function buildItem(item: DataItem, cell: RefResult, indexes: RefResult[]): BuiltItem {
  const read = usable(item, "", cell);
  const indexReads = indexes.map((r) => usable(item, "index ", r));
  const base = { item, ref: read.ref, old: read.values };
  switch (item.type) {
    case "single": {
      if (read.values.length !== 1 || read.values[0].length !== 1) {
        throw new Error(`${item.name}: a Single value must be one cell.`);
      }
      const value = blankToNull(read.values[0][0]);
      return { ...base, value, count: 1, blanks: value === null ? 1 : 0 };
    }
    case "list":
      return { ...base, ...buildList(read) };
    case "indexed":
      return { ...base, ...buildIndexed(item, read, indexReads) };
    case "table":
      return { ...base, ...buildTable(item, read, indexReads[0], indexReads[1]) };
  }
}

export async function readItems(
  excel: ExcelAdapter,
  items: DataItem[],
  sheetId: string,
): Promise<ItemResult[]> {
  const refs = items.flatMap((i) => [i.cellRange, ...i.indexRanges]);
  const results = await excel.readRefs(refs, sheetId);
  let at = 0;
  return items.map((item) => {
    const cell = results[at++];
    const indexes = item.indexRanges.map(() => results[at++]);
    try {
      return buildItem(item, cell, indexes);
    } catch (e) {
      return { item, error: (e as Error).message };
    }
  });
}

export function overlapWarnings(built: BuiltItem[]): string[] {
  const placed = built.filter((b) => b.ref);
  return placed.flatMap((a, i) =>
    placed
      .slice(i + 1)
      .filter((b) => !(a.item.readOnly && b.item.readOnly) && overlaps(a.ref!, b.ref!))
      .map((b) => {
        if (!a.item.readOnly && !b.item.readOnly) {
          return `${a.item.name} and ${b.item.name} share cells, so results can overwrite.`;
        }
        const [locked, other] = a.item.readOnly ? [a, b] : [b, a];
        return `${other.item.name} shares cells with read-only ${locked.item.name}. Those cells will keep their values, so check the range.`;
      }),
  );
}

function repr(v: unknown): string {
  if (v === null) return "None";
  if (v === true) return "True";
  if (v === false) return "False";
  if (typeof v === "string") return `'${v.replace(/\\/g, "\\\\").replace(/'/g, "\\'")}'`;
  if (Array.isArray(v)) {
    return v.length === 1 ? `(${repr(v[0])},)` : `(${v.map(repr).join(", ")})`;
  }
  return String(v);
}

const PREVIEW_ENTRIES = 5;

/** How Python will see the item. */
export function preview(built: BuiltItem): { summary: string; code: string } {
  const { value, count, blanks, item } = built;
  const note = blanks && item.type !== "single" ? `, ${blanks} blank omitted` : "";
  if (Array.isArray(value)) {
    const shown = value.slice(0, PREVIEW_ENTRIES).map(repr);
    if (count > PREVIEW_ENTRIES) shown.push("…");
    return { summary: `list, ${entries(count)}`, code: `[${shown.join(",\n ")}]` };
  }
  if (value !== null && typeof value === "object") {
    const shown = value.__dict__
      .slice(0, PREVIEW_ENTRIES)
      .map(([k, v]) => `${repr(k)}: ${repr(v)}`);
    if (count > PREVIEW_ENTRIES) shown.push("…");
    return { summary: `dict, ${entries(count)}${note}`, code: `{${shown.join(",\n ")}}` };
  }
  const kind = value === null ? "blank" : typeof value === "string" ? "str" : typeof value;
  const pyKind = kind === "boolean" ? "bool" : Number.isInteger(value) ? "int" : kind;
  return { summary: pyKind === "number" ? "float" : pyKind, code: repr(value) };
}

export function toItemIndexes(type: ItemType, index: [string, string]): string[] {
  if (type === "indexed") return splitRefs(index[0]);
  if (type === "table") return index.map((s) => s.trim());
  return [];
}
