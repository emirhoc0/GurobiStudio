// Checks that need no sheet. Problems that depend on cell contents surface when the item is built.
import { parseRef } from "./refs";
import type { DataItem } from "./shapes";

const KEYWORDS = new Set(
  "False None True and as assert async await break class continue def del elif else except finally for from global if import in is lambda nonlocal not or pass raise return try while with yield".split(
    " ",
  ),
);
const IDENTIFIER = /^[\p{L}_][\p{L}\p{N}_]*$/u;
// The runner places this object in the model's globals, so an item of that name would be lost.
const SHIM_NAMES = new Set(["Application"]);

export interface Validation {
  errors: Partial<Record<"name" | "cell" | "index0" | "index1", string>>;
  warnings: string[];
}

function refError(text: string, missing: string): string | undefined {
  if (!text.trim()) return missing;
  try {
    parseRef(text);
    return undefined;
  } catch (e) {
    return (e as Error).message;
  }
}

export function validateItem(
  item: DataItem,
  others: DataItem[],
  reservedNames: string[],
): Validation {
  const errors: Validation["errors"] = {};
  const warnings: string[] = [];
  const { name, type, cellRange, indexRanges } = item;

  if (!name) errors.name = "Give the item a name.";
  else if (!IDENTIFIER.test(name)) {
    errors.name = "Use letters, digits and underscores, not starting with a digit.";
  } else if (KEYWORDS.has(name)) errors.name = `${name} is a Python keyword.`;
  else if (SHIM_NAMES.has(name)) errors.name = `${name} is reserved by the SolverStudio module.`;
  else if (others.some((o) => o.name === name)) errors.name = `${name} is already an item here.`;
  else if (reservedNames.includes(name)) {
    warnings.push(
      `${name} is also a gurobipy name. Importing gurobipy after SolverStudio hides it.`,
    );
  }

  const cell = refError(cellRange, "Enter a cell range.");
  if (cell) errors.cell = cell;
  if (type === "indexed") {
    const bad = indexRanges.map((r) => refError(r, "")).find(Boolean);
    if (!indexRanges.length) errors.index0 = "Enter at least one index range.";
    else if (bad) errors.index0 = bad;
  }
  if (type === "table") {
    const row = refError(indexRanges[0] ?? "", "Enter the row index range.");
    const col = refError(indexRanges[1] ?? "", "Enter the column index range.");
    if (row) errors.index0 = row;
    if (col) errors.index1 = col;
  }
  return { errors, warnings };
}
