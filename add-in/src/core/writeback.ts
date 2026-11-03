// Turns a changed Python value back into cell writes. In Office.js null keeps a cell, so unchanged
// formulas survive, and "" clears it, so leftovers are blanked. Easy to get backwards; see the tests.
import type { Cell, RangeWrite } from "../excel/adapter";
import { columnLetters, rectOf } from "./refs";
import { keyId, type BuiltItem, type Wire } from "./shapes";

const isScalar = (v: unknown): v is Cell | null =>
  v === null || ["string", "number", "boolean"].includes(typeof v);

// Excel parses written strings: "=..." as a formula, "007" as 7. A leading apostrophe stores any
// string as exactly that text. All of them, because a list of risky prefixes is only ever a guess.
const asText = (v: Cell): Cell => (typeof v === "string" && v !== "" ? `'${v}` : v);

export function writeFor(
  built: BuiltItem,
  value: Wire,
): { write?: RangeWrite; warnings: string[] } {
  const { item, ref, old, slots } = built;
  const { name } = item;
  if (!ref) {
    return { warnings: [`${name} changed, but its range has no address, so nothing was written.`] };
  }

  const warnings: string[] = [];
  const next: (Cell | null)[][] = old.map((row) => row.map(() => null));
  let skipped = 0;
  const put = (r: number, c: number, v: unknown) => {
    if (!isScalar(v)) {
      skipped++;
      return;
    }
    const cell = v ?? "";
    if (cell !== old[r][c]) next[r][c] = asText(cell);
  };
  const refuse = (shape: string) => ({
    warnings: [`${name} should hold ${shape}, but the model gave it something else; not written.`],
  });

  if (item.type === "single") {
    if (!isScalar(value)) return refuse("a single value");
    put(0, 0, value);
  } else if (slots) {
    if (value === null || typeof value !== "object" || !("__dict__" in value)) {
      return refuse("a dict");
    }
    const incoming = new Map(value.__dict__.map(([k, v]) => [keyId(k), v]));
    for (const [id, [r, c]] of slots) put(r, c, incoming.get(id) ?? null);
    const outside = [...incoming.keys()].filter((id) => !slots.has(id));
    if (outside.length) {
      warnings.push(
        `${name}: skipped ${outside.length} key(s) not in the index, such as ${outside[0]}.`,
      );
    }
  } else {
    if (!Array.isArray(value)) return refuse("a list");
    const rows = old.length;
    const cols = old[0].length;
    if (rows > 1 && cols > 1) {
      if (value.length > rows) warnings.push(`${name}: cut off after ${rows} rows.`);
      let wide = false;
      for (let r = 0; r < rows; r++) {
        const entry = r < value.length ? value[r] : [];
        const line = Array.isArray(entry) ? entry : [entry];
        wide ||= line.length > cols;
        for (let c = 0; c < cols; c++) put(r, c, c < line.length ? line[c] : null);
      }
      if (wide) warnings.push(`${name}: cut off after ${cols} columns.`);
    } else {
      const length = Math.max(rows, cols);
      if (value.length > length) warnings.push(`${name}: cut off after ${length} entries.`);
      for (let i = 0; i < length; i++) {
        const entry = i < value.length ? value[i] : null;
        const v = Array.isArray(entry) && entry.length === 1 ? entry[0] : entry;
        if (rows > 1) put(i, 0, v);
        else put(0, i, v);
      }
    }
  }

  if (skipped) warnings.push(`${name}: left ${skipped} cell(s) alone that got a non-cell value.`);
  return next.some((row) => row.some((v) => v !== null))
    ? { write: { ref, values: next }, warnings }
    : { warnings };
}

/** Addresses of the cells a mask from formulaCells marks, for showing to the user. */
export function maskedAddresses(writes: RangeWrite[], masks: boolean[][][]): string[] {
  return writes.flatMap((w, i) => {
    const { r1, c1 } = rectOf(w.ref.address);
    return masks[i].flatMap((row, r) =>
      row.flatMap((hit, c) => (hit ? [`${columnLetters(c1 + c)}${r1 + r}`] : [])),
    );
  });
}

/** Leaves masked cells alone, and drops a write with nothing left to do. */
export function keepMasked(writes: RangeWrite[], masks: boolean[][][]): RangeWrite[] {
  return writes
    .map((w, i) => ({
      ...w,
      values: w.values.map((row, r) => row.map((v, c) => (masks[i][r][c] ? null : v))),
    }))
    .filter((w) => w.values.some((row) => row.some((v) => v !== null)));
}
