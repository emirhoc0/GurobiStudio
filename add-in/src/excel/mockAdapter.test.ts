import { describe, expect, it } from "vitest";
import { createMockAdapter } from "./mockAdapter";

const ref = (address: string) => ({ sheetId: "s", address });

describe("mock adapter", () => {
  it("reads blanks as empty strings, like Office.js", async () => {
    const { adapter } = createMockAdapter({ A1: 1, B2: "x" });
    expect(await adapter.readRanges([ref("A1:B2")])).toEqual([
      [
        [1, ""],
        ["", "x"],
      ],
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
});
