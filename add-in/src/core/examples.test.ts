import { describe, expect, it } from "vitest";
import { createMockAdapter } from "../excel/mockAdapter";
import { EXAMPLES, insertExample } from "./examples";
import { freeSheetName } from "./refs";
import { readItems } from "./shapes";

describe("examples", () => {
  it("use every item type between them", () => {
    const types = new Set(EXAMPLES.flatMap((e) => e.items.map((i) => i.type)));
    expect([...types].sort()).toEqual(["indexed", "list", "single", "table"]);
  });

  it.each(EXAMPLES.map((e) => [e.name, e] as const))(
    "%s reads cleanly on its sheet",
    async (_, e) => {
      expect(new Set(e.rows.map((r) => r.length)).size).toBe(1); // Excel needs a rectangle
      const mock = createMockAdapter();
      const sheet = await insertExample(mock.adapter, e);
      expect((await mock.adapter.activeSheet()).id).toBe(sheet.id);
      const results = await readItems(mock.adapter, e.items, sheet.id);
      expect(results.flatMap((r) => ("error" in r ? [r.error] : []))).toEqual([]);
    },
  );

  it("don't reuse a sheet name", async () => {
    const { adapter } = createMockAdapter();
    expect((await insertExample(adapter, EXAMPLES[0])).name).toBe(EXAMPLES[0].name);
    expect((await insertExample(adapter, EXAMPLES[0])).name).toBe(`${EXAMPLES[0].name} 2`);
    expect(freeSheetName("diet", ["DIET", "Diet 2"])).toBe("diet 3");
  });
});
