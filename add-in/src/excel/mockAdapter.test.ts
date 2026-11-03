import { describe, expect, it } from "vitest";
import { createMockAdapter } from "./mockAdapter";

const ref = (address: string) => ({ sheetId: "s", address });

describe("mock adapter", () => {
  it("reads blanks as empty strings, like Office.js", async () => {
    const { adapter } = createMockAdapter({ A1: 1, B2: "x" });
    expect(await adapter.readRefs(["A1:B2"], "s")).toEqual([
      {
        ref: ref("A1:B2"),
        values: [
          [1, ""],
          ["", "x"],
        ],
        errorCells: [],
      },
    ]);
  });

  it("null keeps a cell and an empty string clears it, matching real Excel", async () => {
    const { adapter, cells } = createMockAdapter({ A1: 1, A2: 2, A3: 3 });
    await adapter.writeRanges([{ ref: ref("A1:A3"), values: [[null], [""], [9]] }]);
    expect(Object.fromEntries(cells)).toEqual({ A1: 1, A3: 9 });
  });

  it("handles columns past Z", async () => {
    const { adapter, cells } = createMockAdapter();
    await adapter.writeRanges([{ ref: ref("AA1"), values: [[1]] }]);
    expect(cells.get("AA1")).toBe(1);
  });

  it("re-resolves names and tables on each read, so growth is seen", async () => {
    const mock = createMockAdapter(
      { A2: "a", A3: "b", A4: "c" },
      { tables: { T: { address: "A2:A3", columns: ["Item"] } } },
    );
    const [before] = await mock.adapter.readRefs(["T[Item]"], "s");
    mock.tables.get("T")!.address = "A2:A4";
    const [after] = await mock.adapter.readRefs(["T[Item]"], "s");
    expect(before).toMatchObject({ values: [["a"], ["b"]] });
    expect(after).toMatchObject({ ref: ref("A2:A4"), values: [["a"], ["b"], ["c"]] });
  });

  it("reports an unknown name per reference, not for the whole batch", async () => {
    const { adapter } = createMockAdapter({ A1: 1 });
    const [good, bad] = await adapter.readRefs(["A1", "nope"], "s");
    expect(good).toMatchObject({ values: [[1]] });
    expect(bad).toEqual({ error: "There is no range, name or table called nope." });
  });

  it("gives a formula name values and no address", async () => {
    const { adapter } = createMockAdapter({}, { names: { dyn: [[1], [2]] } });
    expect(await adapter.readRefs(["dyn"], "s")).toEqual([
      { ref: null, values: [[1], [2]], errorCells: [] },
    ]);
  });
});
