// Proves what Excel for Mac actually does. Runs on a temporary sheet, removes everything after,
// and posts the report to the helper so results can be read without copying from the pane.
import { apiFetch } from "./api";

export interface ProbeResult {
  name: string;
  ok: boolean;
  detail?: unknown;
  error?: string;
}

const SHEET = "_gsdiag";
const SHEET2 = "_gsdiag2";
const NAMES = ["gsdiag_plain", "gsdiag_tbl", "gsdiag_off"];
const NS = "urn:gurobistudio:diag";

const escapeXml = (s: string) =>
  s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

async function probe(
  results: ProbeResult[],
  name: string,
  fn: () => Promise<{ ok: boolean; detail?: unknown }>,
) {
  try {
    results.push({ name, ...(await fn()) });
  } catch (e) {
    const err = e as { message?: string; debugInfo?: unknown };
    results.push({ name, ok: false, error: `${err.message}`, detail: err.debugInfo });
  }
}

async function cleanup() {
  await Excel.run(async (ctx) => {
    for (const n of NAMES) ctx.workbook.names.getItemOrNullObject(n).delete();
    for (const s of [SHEET, SHEET2]) ctx.workbook.worksheets.getItemOrNullObject(s).delete();
    await ctx.sync().catch(() => undefined);
    const parts = ctx.workbook.customXmlParts.getByNamespace(NS);
    parts.load("items");
    await ctx.sync();
    parts.items.forEach((p) => p.delete());
    await ctx.sync();
  }).catch(() => undefined);
}

function requirementSets() {
  const check: [string, string][] = [
    ["ExcelApi", "1.5"],
    ["ExcelApi", "1.21"],
    ["SharedRuntime", "1.1"],
    ["SharedRuntime", "1.2"],
    ["DialogApi", "1.2"],
    ["DialogOrigin", "1.1"],
    ["RibbonApi", "1.1"],
  ];
  return Object.fromEntries(
    check.map(([s, v]) => [`${s} ${v}`, Office.context.requirements.isSetSupported(s, v)]),
  );
}

function dialogRoundTrip(): Promise<{ ok: boolean; detail: unknown }> {
  return new Promise((resolve) => {
    const url = `${location.origin}/taskpane.html?view=diag-dialog`;
    Office.context.ui.displayDialogAsync(url, { height: 30, width: 30 }, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        resolve({ ok: false, detail: res.error });
        return;
      }
      const dialog = res.value;
      const timer = setTimeout(() => {
        dialog.close();
        resolve({ ok: false, detail: "timed out waiting for the dialog" });
      }, 10000);
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        const message = (arg as { message?: string }).message;
        if (message === "ready") dialog.messageChild("ping");
        if (message === "pong:ping") {
          clearTimeout(timer);
          dialog.close();
          resolve({ ok: true, detail: message });
        }
      });
    });
  });
}

export function runDialogSide() {
  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
    Office.context.ui.messageParent(`pong:${(arg as { message: string }).message}`);
  });
  Office.context.ui.messageParent("ready");
}

export async function runDiagnostics(): Promise<{ results: ProbeResult[]; saved?: string }> {
  const results: ProbeResult[] = [];
  await cleanup();

  await probe(results, "environment", async () => ({
    ok: true,
    detail: {
      diagnostics: Office.context.diagnostics,
      sets: requirementSets(),
      userAgent: navigator.userAgent,
      officeTheme: Office.context.officeTheme ?? null,
      prefersDark: matchMedia("(prefers-color-scheme: dark)").matches,
    },
  }));

  await probe(results, "custom XML part round trip", () =>
    Excel.run(async (ctx) => {
      const payload = JSON.stringify({
        model: 'from gurobipy import *\n\tprint("<a> & ]]> çğüş ✓")\r\nend',
      });
      ctx.workbook.customXmlParts.add(`<state xmlns="${NS}">${escapeXml(payload)}</state>`);
      await ctx.sync();
      const part = ctx.workbook.customXmlParts.getByNamespace(NS).getOnlyItemOrNullObject();
      const xml = part.getXml();
      await ctx.sync();
      const text = new DOMParser().parseFromString(xml.value, "application/xml").documentElement
        .textContent;
      return { ok: text === payload, detail: { read: text } };
    }),
  );

  await probe(results, "very hidden sheet (persistence fallback)", () =>
    Excel.run(async (ctx) => {
      const s = ctx.workbook.worksheets.add(SHEET2);
      s.visibility = Excel.SheetVisibility.veryHidden;
      await ctx.sync();
      s.load("visibility,id");
      await ctx.sync();
      return { ok: s.visibility === "VeryHidden", detail: { id: s.id } };
    }),
  );

  await probe(results, "null keeps a cell, empty string clears it", () =>
    Excel.run(async (ctx) => {
      const s = ctx.workbook.worksheets.add(SHEET);
      const r = s.getRange("F1:F3");
      r.formulas = [[1], ["=1+1"], [3]];
      await ctx.sync();
      r.values = [[null], [null], [""]] as unknown as string[][];
      await ctx.sync();
      r.load("formulas");
      await ctx.sync();
      const f = r.formulas;
      return { ok: f[0][0] === 1 && f[1][0] === "=1+1" && f[2][0] === "", detail: f };
    }),
  );

  await probe(results, "cell value types", () =>
    Excel.run(async (ctx) => {
      const s = ctx.workbook.worksheets.getItem(SHEET);
      const r = s.getRange("G1:G5");
      r.formulas = [[3], [2.5], ["text"], [true], ["=NA()"]];
      await ctx.sync();
      r.load("values,valueTypes");
      await ctx.sync();
      return {
        ok: r.valueTypes[4][0] === "Error" && r.values[3][0] === true,
        detail: { values: r.values, types: r.valueTypes, json: JSON.stringify(r.values) },
      };
    }),
  );

  await probe(results, "plain named range", () =>
    Excel.run(async (ctx) => {
      const s = ctx.workbook.worksheets.getItem(SHEET);
      ctx.workbook.names.add(NAMES[0], s.getRange("H1:H2"));
      await ctx.sync();
      const r = ctx.workbook.names.getItem(NAMES[0]).getRangeOrNullObject();
      r.load("address");
      await ctx.sync();
      return { ok: !r.isNullObject, detail: r.isNullObject ? null : r.address };
    }),
  );

  await probe(results, "Table reference grows with the table", () =>
    Excel.run(async (ctx) => {
      const s = ctx.workbook.worksheets.getItem(SHEET);
      s.getRange("A1:B3").values = [
        ["Item", "Cost"],
        ["a", 1],
        ["b", 2],
      ];
      const t = s.tables.add("A1:B3", true);
      t.name = "gsdiagT";
      await ctx.sync();
      ctx.workbook.names.add(NAMES[1], "=gsdiagT[Cost]");
      await ctx.sync();

      const named = ctx.workbook.names.getItem(NAMES[1]);
      named.load("type,formula");
      const viaName = named.getRangeOrNullObject();
      viaName.load("address");
      const viaTable = t.columns.getItem("Cost").getDataBodyRange();
      viaTable.load("address");
      await ctx.sync();
      const before = {
        nameType: named.type,
        formula: named.formula,
        viaName: viaName.isNullObject ? null : viaName.address,
        viaTable: viaTable.address,
      };

      t.rows.add(undefined, [["c", 3]]);
      await ctx.sync();
      const viaName2 = named.getRangeOrNullObject();
      viaName2.load("address");
      const viaTable2 = t.columns.getItem("Cost").getDataBodyRange();
      viaTable2.load("address");
      await ctx.sync();
      const after = {
        viaName: viaName2.isNullObject ? null : viaName2.address,
        viaTable: viaTable2.address,
      };
      return { ok: after.viaTable !== before.viaTable, detail: { before, after } };
    }),
  );

  await probe(results, "OFFSET-style dynamic name", () =>
    Excel.run(async (ctx) => {
      ctx.workbook.names.add(NAMES[2], `=OFFSET(${SHEET}!$A$1,0,0,2,1)`);
      await ctx.sync();
      const named = ctx.workbook.names.getItem(NAMES[2]);
      named.load("type,formula,value");
      const r = named.getRangeOrNullObject();
      r.load("address");
      await ctx.sync();
      return {
        ok: !r.isNullObject,
        detail: {
          type: named.type,
          value: named.value,
          address: r.isNullObject ? null : r.address,
        },
      };
    }),
  );

  await probe(results, "worksheet id survives rename", () =>
    Excel.run(async (ctx) => {
      const s = ctx.workbook.worksheets.getItem(SHEET2);
      s.load("id");
      await ctx.sync();
      const before = s.id;
      s.name = `${SHEET2}x`;
      await ctx.sync();
      s.load("id");
      await ctx.sync();
      const after = s.id;
      s.name = SHEET2;
      await ctx.sync();
      return { ok: before === after, detail: { before, after } };
    }),
  );

  await probe(results, "fetch streams SSE incrementally", async () => {
    const t0 = performance.now();
    const res = await apiFetch("/api/diag/stream");
    const reader = res.body!.getReader();
    const arrivals: number[] = [];
    for (;;) {
      const { done } = await reader.read();
      if (done) break;
      arrivals.push(Math.round(performance.now() - t0));
    }
    return {
      ok: arrivals.length >= 3 && arrivals[arrivals.length - 1] - arrivals[0] > 800,
      detail: arrivals,
    };
  });

  await probe(results, "dialog two-way messaging", dialogRoundTrip);

  await cleanup();
  results.push({
    name: "temporary sheet and names removed",
    ...(await Excel.run(async (ctx) => {
      const s = ctx.workbook.worksheets.getItemOrNullObject(SHEET);
      const parts = ctx.workbook.customXmlParts.getByNamespace(NS);
      parts.load("items");
      s.load("name");
      await ctx.sync();
      return { ok: s.isNullObject && parts.items.length === 0 };
    })),
  });

  const res = await apiFetch("/api/diag", { method: "POST", body: JSON.stringify({ results }) });
  return { results, saved: res.ok ? (await res.json()).saved : undefined };
}
