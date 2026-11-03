// Proves what Excel for Mac actually does. Runs on a temporary sheet, removes everything after,
// and posts the report to the helper so results can be read without copying from the pane.
import { officeAdapter } from "../excel/officeAdapter";
import { apiFetch } from "./api";
import { buildItem } from "./shapes";
import { writeFor } from "./writeback";

export interface ProbeResult {
  name: string;
  ok: boolean;
  detail?: unknown;
  error?: string;
}

const SHEET = "_gsdiag";
const SHEET2 = "_gsdiag2";
const NAMES = ["gsdiag_plain", "gsdiag_tbl", "gsdiag_off", "gsdiag_made"];
// Excel's API can't create an OFFSET name, so this one is made by hand and never deleted here.
const USER_OFFSET_NAME = "gs_offset";
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
    await ctx.sync().catch(() => undefined);
    for (const name of [SHEET, SHEET2, `${SHEET2}x`]) {
      const sheet = ctx.workbook.worksheets.getItemOrNullObject(name);
      sheet.load("name");
      await ctx.sync();
      if (sheet.isNullObject) continue;
      // Excel won't delete a very hidden sheet.
      sheet.visibility = Excel.SheetVisibility.visible;
      await ctx.sync();
      sheet.delete();
      await ctx.sync();
    }
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
    const key = new URLSearchParams(location.search).get("k") ?? "";
    const url = `${location.origin}/taskpane.html?k=${encodeURIComponent(key)}&view=diag-dialog`;
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
      ctx.workbook.names.add(NAMES[2], `=OFFSET('${SHEET}'!$A$1,0,0,2,1)`);
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

  const diagSheetId = () =>
    Excel.run(async (ctx) => {
      const s = ctx.workbook.worksheets.getItem(SHEET).load("id");
      await ctx.sync();
      return s.id;
    });

  await probe(results, "adapter reads every kind of reference in one batch", async () => {
    const refs = [
      "F1:F3",
      `'${SHEET}'!G1:G5`,
      "gsdiagT[Cost]",
      "gsdiagT[[Item]:[Cost]]",
      "gsdiagT",
      NAMES[0],
      "no_such_name",
    ];
    const reads = await officeAdapter.readRefs(refs, await diagSheetId());
    const errorCell = reads[1];
    const good = reads.slice(0, -1).every((r) => !("error" in r));
    const ok =
      good &&
      "error" in reads[6] &&
      !("error" in errorCell) &&
      errorCell.errorCells[0]?.[0] === 4 &&
      !("error" in reads[2]) &&
      reads[2].ref?.address === "B2:B4";
    return { ok, detail: reads };
  });

  await probe(results, "adapter reports a bad reference without losing the others", async () => {
    const reads = await officeAdapter.readRefs(["F1", "NoSuchSheet!A1"], await diagSheetId());
    return { ok: !("error" in reads[0]) && "error" in reads[1], detail: reads };
  });

  await probe(results, "adapter writes: null keeps a formula, empty string clears", async () => {
    const sheetId = await diagSheetId();
    await Excel.run(async (ctx) => {
      ctx.workbook.worksheets.getItem(SHEET).getRange("J1:J3").formulas = [["=1+1"], [5], [6]];
      await ctx.sync();
    });
    await officeAdapter.writeRanges([
      { ref: { sheetId, address: "J1:J3" }, values: [[null], [""], [7]] },
    ]);
    return Excel.run(async (ctx) => {
      const r = ctx.workbook.worksheets.getItem(SHEET).getRange("J1:J3").load("formulas");
      await ctx.sync();
      const f = r.formulas;
      return { ok: f[0][0] === "=1+1" && f[1][0] === "" && f[2][0] === 7, detail: f };
    });
  });

  await probe(results, "adapter writes formula-like text as text", async () => {
    const sheetId = await diagSheetId();
    const built = buildItem(
      { name: "x", type: "single", cellRange: "J5", indexRanges: [], readOnly: false },
      { ref: { sheetId, address: "J5" }, values: [[""]], errorCells: [] },
      [],
    );
    await officeAdapter.writeRanges([writeFor(built, "=1+1").write!]);
    return Excel.run(async (ctx) => {
      const r = ctx.workbook.worksheets.getItem(SHEET).getRange("J5");
      r.load("values,formulas,valueTypes");
      await ctx.sync();
      return {
        ok: r.values[0][0] === "=1+1" && r.valueTypes[0][0] === "String",
        detail: { values: r.values, formulas: r.formulas, types: r.valueTypes },
      };
    });
  });

  await probe(results, "adapter creates a name and refuses a duplicate", async () => {
    const ref = { sheetId: await diagSheetId(), address: "J1:J3" };
    await officeAdapter.createName(NAMES[3], ref);
    const [read] = await officeAdapter.readRefs([NAMES[3]], ref.sheetId);
    const refused = await officeAdapter.createName(NAMES[3], ref).then(
      () => false,
      () => true,
    );
    const listed = (await officeAdapter.listNames()).includes(NAMES[3]);
    const ok = !("error" in read) && read.ref?.address === "J1:J3" && refused && listed;
    return { ok, detail: { read, refused, listed } };
  });

  await probe(results, "adapter reads the selection with its sheet", async () => {
    await Excel.run(async (ctx) => {
      const s = ctx.workbook.worksheets.getItem(SHEET);
      s.activate();
      s.getRange("A1:A2").select();
      await ctx.sync();
    });
    const address = await officeAdapter.selection();
    return { ok: address === `${SHEET}!A1:A2`, detail: address };
  });

  await probe(results, "OFFSET name made in Name Manager", () =>
    Excel.run(async (ctx) => {
      const named = ctx.workbook.names.getItemOrNullObject(USER_OFFSET_NAME).load("type,value");
      await ctx.sync();
      if (named.isNullObject) return { ok: true, detail: `no name ${USER_OFFSET_NAME}; skipped` };
      const [read] = await officeAdapter.readRefs([USER_OFFSET_NAME], await diagSheetId());
      return {
        ok: !("error" in read),
        detail: { type: named.type, value: named.value, read },
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
      const sheets = ctx.workbook.worksheets;
      const parts = ctx.workbook.customXmlParts.getByNamespace(NS);
      const names = NAMES.map((n) => ctx.workbook.names.getItemOrNullObject(n));
      parts.load("items");
      sheets.load("items/name");
      names.forEach((n) => n.load("name"));
      await ctx.sync();
      const leftovers = [
        ...sheets.items.map((s) => s.name).filter((n) => n.startsWith("_gsdiag")),
        ...names.filter((n) => !n.isNullObject).map((n) => n.name),
        ...parts.items.map(() => "XML part"),
      ];
      return { ok: leftovers.length === 0, detail: leftovers };
    })),
  });

  const res = await apiFetch("/api/diag", { method: "POST", body: JSON.stringify({ results }) });
  return { results, saved: res.ok ? (await res.json()).saved : undefined };
}
