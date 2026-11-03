import { describe, expect, it } from "vitest";
import {
  buildParams,
  DEFAULT_SETTINGS,
  parseState,
  remapSheets,
  SCHEMA,
  type SheetState,
} from "./state";

const sheet = (name: string): SheetState => ({ name, language: "gurobi", items: [] });

describe("parseState", () => {
  it("round-trips model text exactly", () => {
    const model = 'a\tb\r\nc\n"<&>" ]]> çğüş ✓\n';
    const text = JSON.stringify({
      schema: SCHEMA,
      workbookId: "w",
      settings: DEFAULT_SETTINGS,
      sheets: { s: { ...sheet("A"), model } },
    });
    expect(parseState(text).sheets.s.model).toBe(model);
  });

  it("fills settings added since the workbook was saved", () => {
    const state = parseState(JSON.stringify({ schema: 1, workbookId: "w", settings: {} }));
    expect(state.settings.outputFlag).toBe(true);
    expect(state.sheets).toEqual({});
  });

  it.each(["not json", "null", "{}", '{"schema":1}'])("refuses %s", (text) => {
    expect(() => parseState(text)).toThrow();
  });

  it("refuses a newer schema", () => {
    const text = JSON.stringify({ schema: SCHEMA + 1, workbookId: "w" });
    expect(() => parseState(text)).toThrow(/newer/);
  });
});

describe("remapSheets", () => {
  it("keeps ids when any still match", () => {
    const sheets = { a: sheet("One"), gone: sheet("Two") };
    const out = remapSheets(sheets, [
      { id: "a", name: "One" },
      { id: "new", name: "Two" },
    ]);
    expect(Object.keys(out)).toEqual(["a", "gone"]);
  });

  it("matches by name when no ids match", () => {
    const out = remapSheets({ old1: sheet("One"), old2: sheet("Two") }, [
      { id: "x", name: "Two" },
      { id: "y", name: "One" },
    ]);
    expect(out.y.name).toBe("One");
    expect(out.x.name).toBe("Two");
  });

  it("keeps an entry with no matching name", () => {
    const out = remapSheets({ old: sheet("Gone") }, [{ id: "x", name: "Other" }]);
    expect(Object.keys(out)).toEqual(["old"]);
  });
});

describe("buildParams", () => {
  it("skips blanks and converts numbers", () => {
    expect(buildParams(DEFAULT_SETTINGS)).toEqual({});
    expect(buildParams({ ...DEFAULT_SETTINGS, timeLimit: " 5 ", mipGap: "0.01" })).toEqual({
      TimeLimit: 5,
      MIPGap: 0.01,
    });
  });

  it("only sends OutputFlag when off", () => {
    expect(buildParams({ ...DEFAULT_SETTINGS, outputFlag: false })).toEqual({ OutputFlag: 0 });
  });

  it("lets the table override, keeps text values and skips blank keys", () => {
    const out = buildParams({
      ...DEFAULT_SETTINGS,
      threads: "2",
      custom: [
        { key: "Threads", value: "4" },
        { key: "LogFile", value: "a.log" },
        { key: "", value: "1" },
      ],
    });
    expect(out).toEqual({ Threads: 4, LogFile: "a.log" });
  });
});
