// The bundled examples. Each lands on a new sheet, ready to solve after the usual trust prompt.
import assignment from "../../../examples/assignment.json";
import diet from "../../../examples/diet.json";
import knapsack from "../../../examples/knapsack.json";
import transportation from "../../../examples/transportation.json";
import type { Cell, ExcelAdapter } from "../excel/adapter";
import { itemsStore, modelStore } from "../ui/store";
import type { DataItem } from "./shapes";

export interface Example {
  name: string;
  description: string;
  rows: Cell[][];
  bold: string[];
  items: DataItem[];
  model: string;
  expected: { objective: number; cells: Record<string, number> };
}

export const EXAMPLES = [diet, transportation, knapsack, assignment] as Example[];

export async function insertExample(excel: ExcelAdapter, example: Example) {
  const sheet = await excel.addSheet(example.name, example.rows, example.bold);
  // Before activating, so the editor opens on this model rather than the blank template.
  itemsStore.set({ ...itemsStore.get(), [sheet.id]: example.items });
  modelStore.set({ ...modelStore.get(), [sheet.id]: example.model });
  await excel.activateSheet(sheet.id);
  return sheet;
}
