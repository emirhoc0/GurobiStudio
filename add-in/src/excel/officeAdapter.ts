import { freeSheetName, parseRef } from "../core/refs";
import type { Cell, ExcelAdapter, RangeWrite, RefResult } from "./adapter";

const message = (e: unknown) => (e as Error).message;
const STATE_NS = "urn:gurobistudio:state";
// JSON has no raw newlines, so the text survives XML whitespace handling exactly.
const stateXml = (text: string) =>
  `<state xmlns="${STATE_NS}">${text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")}</state>`;

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

  formulaCells: (writes) =>
    Excel.run(async (ctx) => {
      const ranges = writes.map((w) =>
        ctx.workbook.worksheets.getItem(w.ref.sheetId).getRange(w.ref.address).load("formulas"),
      );
      await ctx.sync();
      return ranges.map((range, i) =>
        (range.formulas as unknown[][]).map((row, r) =>
          row.map(
            (f, c) => writes[i].values[r][c] !== null && typeof f === "string" && f.startsWith("="),
          ),
        ),
      );
    }),

  openEditorDialog: (onMessage, onClosed) =>
    new Promise((resolve, reject) => {
      const url = `${location.origin}${location.pathname}${location.search}&view=editor`;
      Office.context.ui.displayDialogAsync(url, { width: 70, height: 85 }, (res) => {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(res.error.message));
          return;
        }
        const dialog = res.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          onMessage((arg as { message: string }).message);
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => onClosed());
        resolve({ send: (message) => dialog.messageChild(message) });
      });
    }),

  editorChannel: (onMessage) => {
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
      onMessage((arg as { message: string }).message);
    });
    return (message) => Office.context.ui.messageParent(message);
  },

  loadState: () =>
    Excel.run(async (ctx) => {
      const parts = ctx.workbook.customXmlParts.getByNamespace(STATE_NS).load("items");
      await ctx.sync();
      if (!parts.items.length) return null;
      const xml = parts.items[0].getXml();
      await ctx.sync();
      return new DOMParser().parseFromString(xml.value, "application/xml").documentElement
        .textContent;
    }),

  saveState: (text) =>
    Excel.run(async (ctx) => {
      const parts = ctx.workbook.customXmlParts.getByNamespace(STATE_NS).load("items");
      await ctx.sync();
      if (parts.items.length) parts.items[0].setXml(stateXml(text));
      else ctx.workbook.customXmlParts.add(stateXml(text));
      await ctx.sync();
    }),

  listSheets: () =>
    Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets.load("items/id,items/name");
      await ctx.sync();
      return sheets.items.map((s) => ({ id: s.id, name: s.name }));
    }),

  addSheet: (name, rows, bold) =>
    Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets.load("items/name");
      await ctx.sync();
      const free = freeSheetName(
        name,
        sheets.items.map((s) => s.name),
      );
      const sheet = ctx.workbook.worksheets.add(free).load("id,name");
      const range = sheet.getRangeByIndexes(0, 0, rows.length, rows[0].length);
      range.values = rows;
      for (const address of bold) sheet.getRange(address).format.font.bold = true;
      range.format.autofitColumns();
      await ctx.sync();
      return { id: sheet.id, name: sheet.name };
    }),

  activateSheet: (sheetId) =>
    Excel.run(async (ctx) => {
      ctx.workbook.worksheets.getItem(sheetId).activate();
      await ctx.sync();
    }),

  officeDark: () => Office.context.officeTheme?.isDarkTheme ?? null,
};
