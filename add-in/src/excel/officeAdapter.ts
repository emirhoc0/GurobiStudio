import type { Cell, ExcelAdapter, RangeRef, RangeWrite } from "./adapter";

export const officeAdapter: ExcelAdapter = {
  activeSheet: () =>
    Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      sheet.load("id,name");
      await ctx.sync();
      return { id: sheet.id, name: sheet.name };
    }),

  readRanges: (refs: RangeRef[]) =>
    Excel.run(async (ctx) => {
      const ranges = refs.map((r) =>
        ctx.workbook.worksheets.getItem(r.sheetId).getRange(r.address).load("values"),
      );
      await ctx.sync();
      return ranges.map((r) => r.values as Cell[][]);
    }),

  writeRanges: (writes: RangeWrite[]) =>
    Excel.run(async (ctx) => {
      for (const w of writes) {
        const range = ctx.workbook.worksheets.getItem(w.ref.sheetId).getRange(w.ref.address);
        range.values = w.values as Cell[][]; // Office.js treats null as "leave this cell alone"
      }
      await ctx.sync();
    }),
};
