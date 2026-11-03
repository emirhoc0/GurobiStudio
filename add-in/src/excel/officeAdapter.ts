import { parseRef } from "../core/refs";
import type { Cell, ExcelAdapter, RangeWrite, RefResult } from "./adapter";

const message = (e: unknown) => (e as Error).message;

function errorCells(types: string[][]): [number, number][] {
  return types.flatMap((row, r) =>
    row.flatMap((t, c): [number, number][] => (t === "Error" ? [[r, c]] : [])),
  );
}

async function readBatch(refs: string[], sheetId: string): Promise<RefResult[]> {
  return Excel.run(async (ctx) => {
    const { names, tables, worksheets } = ctx.workbook;
    const parsed = refs.map((text) => {
      try {
        return parseRef(text);
      } catch (e) {
        return { kind: "bad" as const, error: message(e) };
      }
    });

    // A sheet-scoped name shadows a workbook one, and a bare table name works like a name.
    const lookups = parsed.map((p) => {
      if (p.kind !== "name") return null;
      const sheet = p.sheet ? worksheets.getItem(p.sheet) : worksheets.getItem(sheetId);
      const scoped = sheet.names.getItemOrNullObject(p.name).load("type,value");
      const global = names.getItemOrNullObject(p.name).load("type,value");
      const table = tables.getItemOrNullObject(p.name).load("name");
      return { scoped, global, table };
    });
    await ctx.sync();

    const found = parsed.map((p, i) => {
      if (p.kind === "bad") return { error: p.error };
      if (p.kind === "a1") {
        const sheet = p.sheet ? worksheets.getItem(p.sheet) : worksheets.getItem(sheetId);
        return { range: sheet.getRange(p.address) };
      }
      if (p.kind === "table") {
        const columns = tables.getItem(p.table).columns;
        const first = columns.getItem(p.first).getDataBodyRange();
        return {
          range:
            p.first === p.last
              ? first
              : first.getBoundingRect(columns.getItem(p.last).getDataBodyRange()),
        };
      }
      const { scoped, global, table } = lookups[i]!;
      const named = scoped.isNullObject ? global : scoped;
      if (!named.isNullObject) return { named, range: named.getRangeOrNullObject() };
      if (!table.isNullObject) return { range: table.getDataBodyRange() };
      return { error: `There is no range, name or table called ${p.name}.` };
    });

    for (const f of found) f.range?.load("address,values,valueTypes,worksheet/id");
    await ctx.sync();

    // A formula name has no range, so values are all Excel will give. Read-only by nature.
    const noRange = found.filter((f) => f.named && f.range!.isNullObject);
    for (const f of noRange) {
      if (f.named!.type === "Array") f.named!.arrayValues.load("values,types");
    }
    if (noRange.length) await ctx.sync();

    return found.map((f, i): RefResult => {
      if (f.error) return { error: f.error };
      if (f.range && !f.range.isNullObject) {
        const address = f.range.address;
        const local = address.slice(address.lastIndexOf("!") + 1);
        if (local.includes(",")) return { error: `${refs[i]} covers several separate areas.` };
        return {
          ref: { sheetId: f.range.worksheet.id, address: local },
          values: f.range.values as Cell[][],
          errorCells: errorCells(f.range.valueTypes),
        };
      }
      const named = f.named!;
      if (named.type === "Array") {
        return {
          ref: null,
          values: named.arrayValues.values as Cell[][],
          errorCells: errorCells(named.arrayValues.types),
        };
      }
      if (["String", "Integer", "Double", "Boolean"].includes(String(named.type))) {
        return { ref: null, values: [[named.value as Cell]], errorCells: [] };
      }
      return { error: `${refs[i]} evaluates to ${String(named.value)}, not to cells.` };
    });
  });
}

export const officeAdapter: ExcelAdapter = {
  activeSheet: () =>
    Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      sheet.load("id,name");
      await ctx.sync();
      return { id: sheet.id, name: sheet.name };
    }),

  onSheetActivated: (listener) => {
    void Excel.run(async (ctx) => {
      ctx.workbook.worksheets.onActivated.add(async () => listener());
      await ctx.sync();
    });
  },

  readRefs: async (refs, sheetId) => {
    try {
      return await readBatch(refs, sheetId);
    } catch {
      // One bad reference fails a whole batch, so retry singly to say which one it was.
      const results: RefResult[] = [];
      for (const text of refs) {
        try {
          results.push(...(await readBatch([text], sheetId)));
        } catch (e) {
          results.push({ error: `Could not read ${text}: ${message(e)}` });
        }
      }
      return results;
    }
  },

  writeRanges: (writes: RangeWrite[]) =>
    Excel.run(async (ctx) => {
      for (const w of writes) {
        const range = ctx.workbook.worksheets.getItem(w.ref.sheetId).getRange(w.ref.address);
        range.values = w.values as Cell[][]; // Office.js treats null as "leave this cell alone"
      }
      await ctx.sync();
    }),

  selection: () =>
    Excel.run(async (ctx) => {
      const areas = ctx.workbook.getSelectedRanges().load("address");
      await ctx.sync();
      return areas.address;
    }),

  listNames: () =>
    Excel.run(async (ctx) => {
      const names = ctx.workbook.names.load("items/name,items/visible");
      const tables = ctx.workbook.tables.load("items/name");
      await ctx.sync();
      return [
        ...names.items.filter((n) => n.visible).map((n) => n.name),
        ...tables.items.map((t) => t.name),
      ];
    }),

  createName: (name, ref) =>
    Excel.run(async (ctx) => {
      const existing = ctx.workbook.names.getItemOrNullObject(name).load("name");
      await ctx.sync();
      if (!existing.isNullObject) throw new Error(`The name ${name} already exists in Excel.`);
      const range = ctx.workbook.worksheets.getItem(ref.sheetId).getRange(ref.address);
      ctx.workbook.names.add(name, range);
      await ctx.sync();
    }),
};
