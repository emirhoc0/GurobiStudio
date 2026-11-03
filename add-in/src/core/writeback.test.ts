import { describe, expect, it } from "vitest";
import type { Cell } from "../excel/adapter";
import { createMockAdapter } from "../excel/mockAdapter";
import { buildItem, type BuiltItem, type DataItem, type ItemType, type Wire } from "./shapes";
import { writeFor } from "./writeback";

const def = (type: ItemType): DataItem => ({
  name: "x",
  type,
  cellRange: "",
  indexRanges: [],
  readOnly: false,
});
const read = (address: string, values: Cell[][]) => ({
  ref: { sheetId: "s", address },
  values,
  errorCells: [] as [number, number][],
});
const list = (address: string, values: Cell[][]) =>
  buildItem(def("list"), read(address, values), []);
const indexed = (values: Cell[][]) =>
  buildItem(def("indexed"), read("B1:B3", values), [read("A1:A3", [["a"], ["b"], ["c"]])]);
const d = (entries: [unknown, unknown][]) => ({ __dict__: entries }) as Wire;

describe("null keeps, empty string clears", () => {
  it("sends null for an unchanged cell, so its formula survives", () => {
    const { write } = writeFor(list("A1:A3", [[1], [2], [3]]), [1, 9, 3]);
    expect(write!.values).toEqual([[null], [9], [null]]);
  });

  it("clears leftovers when a result is shorter than its range", () => {
    const { write } = writeFor(list("A1:A3", [[1], [2], [3]]), [1]);
    expect(write!.values).toEqual([[null], [""], [""]]);
  });

  it("does not touch cells that are already blank", () => {
    const { write } = writeFor(list("A1:A3", [[""], [""], [""]]), [7]);
    expect(write!.values).toEqual([[7], [null], [null]]);
  });

  it("clears a cell when Python assigns None", () => {
    const b = buildItem(def("single"), read("A1", [[5]]), []);
    expect(writeFor(b, null).write!.values).toEqual([[""]]);
  });

  it("writes nothing when no cell differs", () => {
    expect(writeFor(list("A1:A2", [[1], [2]]), [1, 2]).write).toBeUndefined();
  });

  it("writes to the addresses it read", () => {
    expect(writeFor(list("D4:D5", [[1], [2]]), [9, 9]).write!.ref).toEqual({
      sheetId: "s",
      address: "D4:D5",
    });
  });

  it("does what it claims through the adapter", async () => {
    const { adapter, cells } = createMockAdapter({ A1: 1, A2: 2, A3: 3 });
    const { write } = writeFor(list("A1:A3", [[1], [2], [3]]), [1, 9]);
    await adapter.writeRanges([write!]);
    expect(Object.fromEntries(cells)).toEqual({ A1: 1, A2: 9 });
  });
});

describe("text that looks like a formula", () => {
  const one = (old: Cell) => buildItem(def("single"), read("A1", [[old]]), []);

  it.each(["=HYPERLINK(1)", "+1", "-abc", "@x"])("stores %s as text", (text) => {
    expect(writeFor(one(""), text).write!.values).toEqual([[`'${text}`]]);
  });

  it("leaves ordinary values as they are", () => {
    expect(writeFor(one(""), "a=b").write!.values).toEqual([["a=b"]]);
    expect(writeFor(one(""), -5).write!.values).toEqual([[-5]]);
  });

  it("sees text already in the cell as unchanged", () => {
    expect(writeFor(one("=x"), "=x").write).toBeUndefined();
  });

  it("reads back as the same text through the adapter", async () => {
    const { adapter, cells } = createMockAdapter();
    await adapter.writeRanges([writeFor(one(""), "=1+1").write!]);
    expect(cells.get("A1")).toBe("=1+1");
  });

  it("applies to dict and list items too", () => {
    expect(writeFor(list("A1", [[""]]), ["=x"]).write!.values).toEqual([["'=x"]]);
    expect(writeFor(indexed([[""], [""], [""]]), d([["a", "=x"]])).write!.values[0]).toEqual([
      "'=x",
    ]);
  });
});

describe("single", () => {
  it("refuses a list", () => {
    const b = buildItem(def("single"), read("A1", [[5]]), []);
    const result = writeFor(b, [1, 2]);
    expect(result.write).toBeUndefined();
    expect(result.warnings[0]).toContain("single value");
  });
});

describe("list", () => {
  it("writes across a row", () => {
    expect(writeFor(list("A1:C1", [["", "", ""]]), [1, 2]).write!.values).toEqual([[1, 2, null]]);
  });

  it("cuts off a result longer than its range, with a warning", () => {
    const result = writeFor(list("A1:A2", [[""], [""]]), [1, 2, 3]);
    expect(result.write!.values).toEqual([[1], [2]]);
    expect(result.warnings[0]).toContain("cut off after 2");
  });

  it("writes rows of tuples into several columns and clears short rows", () => {
    const b = list("A1:B2", [
      [1, 2],
      [3, 4],
    ]);
    expect(writeFor(b, [[9, 2], [3]]).write!.values).toEqual([
      [9, null],
      [null, ""],
    ]);
  });

  it("warns about tuples wider than the range", () => {
    const b = list("A1:B2", [
      ["", ""],
      ["", ""],
    ]);
    const result = writeFor(b, [[1, 2, 3]]);
    expect(result.write!.values[0]).toEqual([1, 2]);
    expect(result.warnings[0]).toContain("cut off after 2 columns");
  });

  it("refuses a dict", () => {
    expect(writeFor(list("A1:A2", [[1], [2]]), d([])).warnings[0]).toContain("a list");
  });
});

describe("dict items", () => {
  it("writes to keys whose cell started blank", () => {
    const b = indexed([[""], [""], [""]]);
    const { write } = writeFor(b, d([["a", 1], ["c", 3]])); // prettier-ignore
    expect(write!.values).toEqual([[1], [null], [3]]);
  });

  it("blanks a key the model removed", () => {
    const b = indexed([[1], [2], [3]]);
    expect(writeFor(b, d([["a", 1], ["b", 2]])).write!.values).toEqual([[null], [null], [""]]); // prettier-ignore
  });

  it("skips keys outside the index, with a warning", () => {
    const result = writeFor(indexed([[""], [""], [""]]), d([["a", 1], ["z", 9]])); // prettier-ignore
    expect(result.write!.values).toEqual([[1], [null], [null]]);
    expect(result.warnings[0]).toContain('skipped 1 key(s) not in the index, such as "z"');
  });

  it("matches tuple keys for tables", () => {
    const b = buildItem(def("table"), read("B2:C2", [["", ""]]), [
      read("A2", [["w"]]),
      read("B1:C1", [["x", "y"]]),
    ]);
    expect(writeFor(b, d([[["w", "y"], 5]])).write!.values).toEqual([[null, 5]]);
  });

  it("does not confuse the number 1 with the text 1", () => {
    const b = buildItem(def("indexed"), read("B1", [[""]]), [read("A1", [[1]])]);
    expect(writeFor(b, d([["1", 5]])).write).toBeUndefined();
    expect(writeFor(b, d([[1, 5]])).write!.values).toEqual([[5]]);
  });

  it("refuses a scalar", () => {
    expect(writeFor(indexed([[1], [2], [3]]), 5).warnings[0]).toContain("a dict");
  });

  it("leaves a cell alone when its value is not a scalar", () => {
    const value = d([["a", [1, 2]]]);
    const result = writeFor(indexed([[1], [""], [""]]), value);
    expect(result.write).toBeUndefined();
    expect(result.warnings[0]).toContain("left 1 cell(s) alone");
  });
});

describe("no address", () => {
  it("writes nothing for a formula name and says so", () => {
    const b: BuiltItem = buildItem(def("list"), { ref: null, values: [[1]], errorCells: [] }, []);
    const result = writeFor(b, [2]);
    expect(result.write).toBeUndefined();
    expect(result.warnings[0]).toContain("no address");
  });
});
