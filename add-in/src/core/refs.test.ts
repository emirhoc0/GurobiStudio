import { describe, expect, it } from "vitest";
import { cellAt, localize, overlaps, parseRef, rowChunks, splitRefs } from "./refs";
import { validateItem } from "./validate";
import type { DataItem } from "./shapes";

describe("parseRef", () => {
  it.each([
    ["A1", { kind: "a1", address: "A1" }],
    ["a1:b3", { kind: "a1", address: "A1:B3" }],
    ["$A$1:$B$3", { kind: "a1", address: "A1:B3" }],
    ["B3:A1", { kind: "a1", address: "A1:B3" }],
    ["A1:A1", { kind: "a1", address: "A1" }],
    ["XFD1048576", { kind: "a1", address: "XFD1048576" }],
    ["Sheet2!A1:B2", { kind: "a1", sheet: "Sheet2", address: "A1:B2" }],
    ["'My Sheet'!A1", { kind: "a1", sheet: "My Sheet", address: "A1" }],
    ["'It''s'!A1", { kind: "a1", sheet: "It's", address: "A1" }],
    ["=Sheet2!$A$1", { kind: "a1", sheet: "Sheet2", address: "A1" }],
    ["supply", { kind: "name", name: "supply" }],
    ["Sheet2!supply", { kind: "name", sheet: "Sheet2", name: "supply" }],
    ["ZZZ1", { kind: "name", name: "ZZZ1" }], // past the last column, so Excel allows it as a name
    ["A0", { kind: "name", name: "A0" }],
    ["Table1[Cost]", { kind: "table", table: "Table1", first: "Cost", last: "Cost" }],
    [
      "Table1[Unit Cost]",
      { kind: "table", table: "Table1", first: "Unit Cost", last: "Unit Cost" },
    ],
    ["Table1[[Cost]]", { kind: "table", table: "Table1", first: "Cost", last: "Cost" }],
    ["Table1[[From]:[To]]", { kind: "table", table: "Table1", first: "From", last: "To" }],
  ])("reads %s", (text, expected) => {
    expect(parseRef(text)).toEqual(expected);
  });

  it.each([
    ["A:A", "whole row or column"],
    ["1:5", "whole row or column"],
    ["Sheet2!A:C", "whole row or column"],
    ["Table1[#All]", "not supported"],
    ["A1,B2", "not a cell range"],
    ["1abc", "not a cell range"],
    ["a b", "not a cell range"],
  ])("rejects %s", (text, message) => {
    expect(() => parseRef(text)).toThrow(message);
  });
});

describe("helpers", () => {
  it("splits on commas outside quotes", () => {
    expect(splitRefs("A1:A3, 'a,b'!C1:C3,")).toEqual(["A1:A3", "'a,b'!C1:C3"]);
    expect(splitRefs("  ")).toEqual([]);
  });

  it("drops the model sheet from a selection, and keeps other sheets", () => {
    expect(localize("Sheet1!A1:A3", "Sheet1")).toBe("A1:A3");
    expect(localize("Sheet2!A1:A3", "Sheet1")).toBe("Sheet2!A1:A3");
    expect(localize("'My Sheet'!A1,'My Sheet'!C1", "My Sheet")).toBe("A1, C1");
  });

  it("names a cell inside a range, past column Z", () => {
    expect(cellAt("Y2:AB9", 1, 2)).toBe("AA3");
  });

  it("detects overlap only on the same sheet", () => {
    const at = (sheetId: string, address: string) => ({ sheetId, address });
    expect(overlaps(at("s", "A1:B2"), at("s", "B2:C3"))).toBe(true);
    expect(overlaps(at("s", "A1:B2"), at("s", "C1:C3"))).toBe(false);
    expect(overlaps(at("s", "A1:B2"), at("t", "A1:B2"))).toBe(false);
  });
});

describe("validateItem", () => {
  const item = (over: Partial<DataItem> = {}): DataItem => ({
    name: "supply",
    type: "list",
    cellRange: "A1:A3",
    indexRanges: [],
    readOnly: false,
    ...over,
  });
  const errors = (over: Partial<DataItem>, others: DataItem[] = []) =>
    validateItem(item(over), others, []).errors;

  it("accepts a good item", () => {
    expect(errors({})).toEqual({});
    expect(errors({ name: "çözüm_1" })).toEqual({});
  });

  it.each([
    ["", "Give the item"],
    ["1x", "not starting with a digit"],
    ["a-b", "not starting with a digit"],
    ["for", "keyword"],
    ["Application", "reserved"],
  ])("rejects the name %s", (name, message) => {
    expect(errors({ name }).name).toContain(message);
  });

  it("rejects a duplicate name", () => {
    expect(errors({}, [item()]).name).toContain("already");
  });

  it("warns on a gurobipy name without blocking", () => {
    const result = validateItem(item({ name: "quicksum" }), [], ["quicksum"]);
    expect(result.errors).toEqual({});
    expect(result.warnings[0]).toContain("gurobipy");
  });

  it("needs the ranges each type uses", () => {
    expect(errors({ cellRange: "" }).cell).toContain("Enter");
    expect(errors({ cellRange: "A:A" }).cell).toContain("whole row");
    expect(errors({ type: "indexed" }).index0).toContain("at least one");
    expect(errors({ type: "indexed", indexRanges: ["A1", "b c"] }).index0).toContain("not a cell");
    expect(errors({ type: "table", indexRanges: ["A1", ""] }).index1).toContain("column index");
    expect(errors({ type: "table", indexRanges: ["", "A1"] }).index0).toContain("row index");
  });
});

describe("rowChunks", () => {
  it("covers every row once, within the cell limit", () => {
    expect(rowChunks(10, 3, 9)).toEqual([
      { start: 0, count: 3 },
      { start: 3, count: 3 },
      { start: 6, count: 3 },
      { start: 9, count: 1 },
    ]);
  });

  it("keeps a small range whole", () => {
    expect(rowChunks(2, 2, 100)).toEqual([{ start: 0, count: 2 }]);
  });

  it("takes at least one row, even when a row is over the limit", () => {
    expect(rowChunks(2, 50, 10)).toEqual([
      { start: 0, count: 1 },
      { start: 1, count: 1 },
    ]);
  });
});
