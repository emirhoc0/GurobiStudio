import { describe, expect, it } from "vitest";
import type { Cell, RefResult } from "../excel/adapter";
import { createMockAdapter } from "../excel/mockAdapter";
import {
  buildItem,
  overlapWarnings,
  preview,
  readItems,
  type BuiltItem,
  type DataItem,
  type ItemType,
} from "./shapes";

const read = (address: string, values: Cell[][]): RefResult => ({
  ref: { sheetId: "s", address },
  values,
  errorCells: [],
});
const def = (type: ItemType, cellRange = "A1", indexRanges: string[] = []): DataItem => ({
  name: "x",
  type,
  cellRange,
  indexRanges,
  readOnly: false,
});
const dict = (b: BuiltItem) => (b.value as { __dict__: unknown[] }).__dict__;

describe("single", () => {
  it("is a scalar, and a blank is None", () => {
    expect(buildItem(def("single"), read("A1", [[5]]), []).value).toBe(5);
    expect(buildItem(def("single"), read("A1", [[""]]), []).value).toBeNull();
  });

  it("rejects several cells", () => {
    expect(() => buildItem(def("single"), read("A1:A2", [[1], [2]]), [])).toThrow("one cell");
  });
});

describe("list", () => {
  it("reads a column or a row as a flat list, ignoring trailing blanks", () => {
    expect(buildItem(def("list"), read("A1:A4", [["a"], ["b"], [""], [""]]), []).value).toEqual([
      "a",
      "b",
    ]);
    expect(buildItem(def("list"), read("A1:D1", [[1, 2, 3, ""]]), []).value).toEqual([1, 2, 3]);
  });

  it("keeps a blank inside the list as None, so positions hold", () => {
    expect(buildItem(def("list"), read("A1:A3", [["a"], [""], ["c"]]), []).value).toEqual([
      "a",
      null,
      "c",
    ]);
  });

  it("reads several columns as rows of tuples", () => {
    const b = buildItem(
      def("list"),
      read("A1:B3", [
        ["a", "x"],
        ["b", "y"],
        ["", ""],
      ]),
      [],
    );
    expect(b.value).toEqual([
      ["a", "x"],
      ["b", "y"],
    ]);
  });
});

describe("indexed", () => {
  const index = read("A1:A3", [["a"], ["b"], ["c"]]);

  it("makes a dict and omits blank values, counting them", () => {
    const b = buildItem(def("indexed"), read("B1:B3", [[1], [""], [3]]), [index]);
    expect(dict(b)).toEqual([
      ["a", 1],
      ["c", 3],
    ]);
    expect(b.blanks).toBe(1);
  });

  it("still knows every key's cell when the output column is empty", () => {
    const b = buildItem(def("indexed"), read("B1:B3", [[""], [""], [""]]), [index]);
    expect(dict(b)).toEqual([]);
    expect([...b.slots!.keys()]).toEqual(['"a"', '"b"', '"c"']);
  });

  it("gives tuple keys from a multi-column index", () => {
    const multi = read("A1:B2", [
      ["a", 1],
      ["b", 2],
    ]);
    expect(dict(buildItem(def("indexed"), read("C1:C2", [[5], [6]]), [multi]))).toEqual([
      [["a", 1], 5],
      [["b", 2], 6],
    ]);
  });

  it("gives tuple keys from several index ranges", () => {
    const b = buildItem(def("indexed"), read("C1:C2", [[5], [6]]), [
      read("A1:A2", [["a"], ["b"]]),
      read("E1:E2", [[1], [2]]),
    ]);
    expect(dict(b)[1]).toEqual([["b", 2], 6]);
  });

  it("works across a row", () => {
    const b = buildItem(def("indexed"), read("A2:C2", [[1, 2, 3]]), [
      read("A1:C1", [["a", "b", "c"]]),
    ]);
    expect(dict(b)[2]).toEqual(["c", 3]);
    expect(b.slots!.get('"c"')).toEqual([0, 2]);
  });

  it("ignores trailing blanks in the index and cell range", () => {
    const b = buildItem(def("indexed"), read("B1:B9", [[1], [2], ...Array(7).fill([""])]), [
      read("A1:A4", [["a"], ["b"], [""], [""]]),
    ]);
    expect(dict(b)).toHaveLength(2);
  });

  it.each<[string, RefResult, RefResult, string]>([
    ["short data", read("B1:B2", [[1], [2]]), index, "only fits 2"],
    ["data past the index", read("B1:B4", [[1], [2], [3], [4]]), index, "past the end"],
    ["a grid", read("B1:C3", [[1, 1], [2, 2], [3, 3]]), index, "Use Table"], // prettier-ignore
    ["an empty index", read("B1:B2", [[1], [2]]), read("A1:A2", [[""], [""]]), "empty"],
    ["a gap in the index", read("B1:B3", [[1], [2], [3]]), read("A1:A3", [["a"], [""], ["c"]]), "blank"], // prettier-ignore
    ["a repeated key", read("B1:B2", [[1], [2]]), read("A1:A2", [["a"], ["a"]]), 'repeats "a"'],
    ["crossed directions", read("B1:B3", [[1], [2], [3]]), read("A1:C1", [["a", "b", "c"]]), "different directions"], // prettier-ignore
  ])("rejects %s", (_, cell, idx, message) => {
    expect(() => buildItem(def("indexed"), cell, [idx])).toThrow(message);
  });
});

describe("table", () => {
  const rows = read("A2:A3", [["w1"], ["w2"]]);
  const cols = read("B1:C1", [["b1", "b2"]]);

  it("makes (row, col) keys and omits blanks for sparse tables", () => {
    const b = buildItem(
      def("table"),
      read("B2:C3", [
        [1, ""],
        [3, 4],
      ]),
      [rows, cols],
    );
    expect(dict(b)).toEqual([
      [["w1", "b1"], 1],
      [["w2", "b1"], 3],
      [["w2", "b2"], 4],
    ]);
    expect(b.blanks).toBe(1);
    expect(b.slots!.get('["w1","b2"]')).toEqual([0, 1]);
  });

  it("rejects a cell range that does not fit the indexes", () => {
    expect(() => buildItem(def("table"), read("B2:B3", [[1], [2]]), [rows, cols])).toThrow(
      "only fits 1",
    );
  });

  it("rejects a two-dimensional row index", () => {
    const wide = read("A2:B3", [["a", "b"], ["c", "d"]]); // prettier-ignore
    expect(() => buildItem(def("table"), read("C2:D3", [[1, 1], [1, 1]]), [wide, cols])).toThrow(
      "single row or column",
    ); // prettier-ignore
  });
});

describe("error cells", () => {
  it("stop the build and name the cell", () => {
    const cell: RefResult = {
      ref: { sheetId: "s", address: "B2:B3" },
      values: [[1], ["#N/A"]],
      errorCells: [[1, 0]],
    };
    expect(() => buildItem(def("list"), cell, [])).toThrow("cell B3 holds #N/A");
  });
});

describe("readItems", () => {
  it("resolves every item in one adapter call and reports errors per item", async () => {
    const mock = createMockAdapter({ A1: "a", A2: "b", B1: 1, B2: 2, D1: 9 });
    let calls = 0;
    const counted = {
      ...mock.adapter,
      readRefs: (refs: string[], id: string) => {
        calls++;
        return mock.adapter.readRefs(refs, id);
      },
    };
    const [good, bad, single] = await readItems(
      counted,
      [
        def("indexed", "B1:B2", ["A1:A2"]),
        { ...def("list", "missing"), name: "y" },
        { ...def("single", "D1"), name: "z" },
      ],
      "s",
    );
    expect(calls).toBe(1);
    expect(dict(good as BuiltItem)).toEqual([
      ["a", 1],
      ["b", 2],
    ]);
    expect(bad).toMatchObject({ error: "y: There is no range, name or table called missing." });
    expect((single as BuiltItem).value).toBe(9);
  });
});

describe("overlapWarnings", () => {
  const at = (name: string, address: string, readOnly = false) =>
    ({ item: { ...def("list"), name, readOnly }, ref: { sheetId: "s", address } }) as BuiltItem;

  it("warns on writable items that share cells, and not otherwise", () => {
    expect(overlapWarnings([at("a", "A1:A5"), at("b", "A5:B9")])).toHaveLength(1);
    expect(overlapWarnings([at("a", "A1:A4"), at("b", "A5:B9")])).toEqual([]);
    expect(overlapWarnings([at("a", "A1:A5"), at("b", "A5:B9", true)])).toEqual([]);
  });
});

describe("preview", () => {
  it("shows Python syntax", () => {
    const t = buildItem(def("single"), read("A1", [[true]]), []);
    expect(preview(t)).toEqual({ summary: "bool", code: "True" });
    expect(preview(buildItem(def("single"), read("A1", [[2.5]]), []))).toEqual({
      summary: "float",
      code: "2.5",
    });
    expect(preview(buildItem(def("single"), read("A1", [[3]]), [])).summary).toBe("int");
    const pairs = buildItem(def("list"), read("A1:B1", [["it's", 1]]), []);
    expect(preview(pairs).code).toBe("['it\\'s',\n 1]");
  });

  it("shows tuple keys and notes omitted blanks", () => {
    const b = buildItem(def("table"), read("B2:C2", [[1, ""]]), [
      read("A2", [["w"]]),
      read("B1:C1", [["x", "y"]]),
    ]);
    expect(preview(b)).toEqual({
      summary: "dict, 1 entry, 1 blank omitted",
      code: "{('w', 'x'): 1}",
    });
  });

  it("cuts long values off", () => {
    const b = buildItem(
      def("list"),
      read(
        "A1:A7",
        Array.from({ length: 7 }, (_, i) => [i]),
      ),
      [],
    );
    expect(preview(b).code.split("\n")).toHaveLength(6);
    expect(preview(b).summary).toBe("list, 7 entries");
  });
});
