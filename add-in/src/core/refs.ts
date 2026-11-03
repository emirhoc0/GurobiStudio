// What a user can type in a range field. Refs are stored as typed and resolved at run time.
import type { RangeRef } from "../excel/adapter";

export type ParsedRef =
  | { kind: "a1"; sheet?: string; address: string }
  | { kind: "table"; table: string; first: string; last: string }
  | { kind: "name"; sheet?: string; name: string };

export interface Rect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

const MAX_ROW = 1048576;
const MAX_COL = 16384;
const NAME = /^[\p{L}_\\][\p{L}\p{N}_.\\?]*$/u;

export function columnNumber(letters: string): number {
  return [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
}

export function columnLetters(n: number): string {
  return n === 0
    ? ""
    : columnLetters(Math.floor((n - 1) / 26)) + String.fromCharCode(65 + ((n - 1) % 26));
}

function parseA1(text: string): string | null {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(text);
  if (!match) return null;
  const [, a, b, c = a, d = b] = match;
  const [c1, c2] = [a, c].map((s) => columnNumber(s.toUpperCase()));
  const [r1, r2] = [b, d].map(Number);
  if (Math.max(c1, c2) > MAX_COL || Math.max(r1, r2) > MAX_ROW || Math.min(r1, r2) < 1) return null;
  const first = columnLetters(Math.min(c1, c2)) + Math.min(r1, r2);
  const last = columnLetters(Math.max(c1, c2)) + Math.max(r1, r2);
  return first === last ? first : `${first}:${last}`;
}

function parseTable(text: string): ParsedRef | null {
  const match = /^([\p{L}_\\][\p{L}\p{N}_.\\]*)\[(.+)\]$/u.exec(text);
  if (!match) return null;
  const [, table, inner] = match;
  if (inner.includes("#")) {
    throw new Error(`Use a column such as ${table}[Cost]; ${text} is not supported.`);
  }
  const span = /^\[([^\]]+)\]:\[([^\]]+)\]$/.exec(inner);
  if (span) return { kind: "table", table, first: span[1], last: span[2] };
  const column = inner.replace(/^\[(.*)\]$/, "$1");
  return { kind: "table", table, first: column, last: column };
}

/** Throws with a message fit to show under the field. */
export function parseRef(input: string): ParsedRef {
  const text = input.trim().replace(/^=/, "");
  const sheetMatch = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/.exec(text);
  const sheet = sheetMatch && (sheetMatch[1]?.replace(/''/g, "'") ?? sheetMatch[2]);
  const rest = sheetMatch ? sheetMatch[3] : text;

  if (/^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$|^\$?\d+:\$?\d+$/.test(rest)) {
    throw new Error(`${rest} is a whole row or column. Use a bounded range or a named range.`);
  }
  const address = parseA1(rest);
  if (address) return { kind: "a1", ...(sheet ? { sheet } : {}), address };
  if (!sheet) {
    const table = parseTable(rest);
    if (table) return table;
  }
  if (NAME.test(rest)) return { kind: "name", ...(sheet ? { sheet } : {}), name: rest };
  throw new Error(`"${input.trim()}" is not a cell range, name or table column.`);
}

/** Splits "A1:A3, 'a,b'!C1:C3" on the commas that separate ranges. */
export function splitRefs(text: string): string[] {
  const parts: string[] = [];
  let current = "";
  let quoted = false;
  for (const ch of text) {
    if (ch === "'") quoted = !quoted;
    if (ch === "," && !quoted) {
      parts.push(current);
      current = "";
    } else {
      current += ch;
    }
  }
  parts.push(current);
  return parts.map((p) => p.trim()).filter(Boolean);
}

export function rectOf(address: string): Rect {
  const [, a, b, c = a, d = b] = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(address) ?? [];
  if (!a) throw new Error(`Bad address: ${address}`);
  return { r1: Number(b), c1: columnNumber(a), r2: Number(d), c2: columnNumber(c) };
}

export function cellAt(address: string, row: number, col: number): string {
  const { r1, c1 } = rectOf(address);
  return columnLetters(c1 + col) + (r1 + row);
}

export function overlaps(a: RangeRef, b: RangeRef): boolean {
  if (a.sheetId !== b.sheetId) return false;
  const [x, y] = [rectOf(a.address), rectOf(b.address)];
  return x.r1 <= y.r2 && y.r1 <= x.r2 && x.c1 <= y.c2 && y.c1 <= x.c2;
}

export function quoteSheet(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

/** Excel's selection address always names the sheet; drop that when it is the model's own sheet. */
export function localize(address: string, sheetName: string): string {
  return splitRefs(address)
    .map((part) => {
      for (const prefix of [`${sheetName}!`, `${quoteSheet(sheetName)}!`]) {
        if (part.startsWith(prefix)) return part.slice(prefix.length);
      }
      return part;
    })
    .join(", ");
}

/** Excel compares sheet names ignoring case. */
export function freeSheetName(name: string, taken: string[]): string {
  const used = new Set(taken.map((t) => t.toLowerCase()));
  let free = name;
  for (let i = 2; used.has(free.toLowerCase()); i++) free = `${name} ${i}`;
  return free;
}
