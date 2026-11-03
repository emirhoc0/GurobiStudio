import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { ExcelAdapter, RefResult } from "../excel/adapter";

vi.stubGlobal("document", { querySelector: () => null }); // api.ts reads its token at load

const { createMockAdapter } = await import("../excel/mockAdapter");
const { solve, stop } = await import("./solve");
const { confirmStore, itemsStore, runStore } = await import("../ui/store");
const { loadWorkbook } = await import("./persist");

const encoder = new TextEncoder();
const sse = (...events: unknown[]) => events.map((e) => `data: ${JSON.stringify(e)}\n\n`).join("");
const STOPPED = { type: "status", state: "stopped" };
const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });
const stream = (...chunks: string[]) =>
  new Response(
    new ReadableStream({
      start(controller) {
        chunks.forEach((c) => controller.enqueue(encoder.encode(c)));
        controller.close();
      },
    }),
  );

type Route = (init?: RequestInit) => Response | Promise<Response>;
let calls: string[];
function route(routes: Record<string, Route>) {
  vi.stubGlobal(
    "fetch",
    vi.fn(async (path: string, init?: RequestInit) => {
      calls.push(`${init?.method ?? "GET"} ${path}`);
      if (path === "/api/health") return json({});
      const handler = routes[path];
      if (!handler) throw new Error(`Unexpected ${path}`);
      return handler(init);
    }),
  );
}
const answerConfirm = (answer: boolean) => {
  const unsubscribe = confirmStore.subscribe(() => {
    confirmStore.get()?.resolve(answer);
    unsubscribe();
  });
};

const SHEET = "mock-sheet";
const excelWith = (cells = {}) => createMockAdapter(cells).adapter;

beforeEach(async () => {
  calls = [];
  runStore.set({ phase: "idle", log: [] });
  itemsStore.set({});
  await loadWorkbook(excelWith());
});
afterEach(() => vi.unstubAllGlobals());
vi.stubGlobal("crypto", globalThis.crypto);

describe("Stop before the run exists", () => {
  it("never starts a run when Stop came during the read", async () => {
    route({});
    let release: (r: RefResult[]) => void = () => {};
    const excel: ExcelAdapter = {
      ...excelWith(),
      readRefs: () => new Promise((resolve) => (release = resolve)),
    };
    const solving = solve(excel);
    await vi.waitFor(() => expect(calls).toContain("GET /api/health"));
    await new Promise((r) => setTimeout(r));
    await stop();
    release([]);
    await solving;
    expect(runStore.get().phase).toBe("stopped");
    expect(calls).not.toContain("POST /api/run");
  });

  it("cancels the run when Stop came while it was starting", async () => {
    let started: (r: Response) => void = () => {};
    route({
      "/api/run": () => new Promise((resolve) => (started = resolve)),
      "/api/run/r1/cancel": () => json({ ok: true }),
      "/api/run/r1/events": () => stream(sse(STOPPED)),
    });
    const solving = solve(excelWith());
    await vi.waitFor(() => expect(calls).toContain("POST /api/run"));
    await stop();
    started(json({ run_id: "r1" }));
    await solving;
    expect(calls).toContain("POST /api/run/r1/cancel");
    expect(runStore.get().phase).toBe("stopped");
  });

  it("forgets a Stop from an earlier solve", async () => {
    route({ "/api/run": () => json({ error: "x" }, 500) });
    let release: (r: RefResult[]) => void = () => {};
    const excel: ExcelAdapter = {
      ...excelWith(),
      readRefs: () => new Promise((resolve) => (release = resolve)),
    };
    const first = solve(excel);
    await vi.waitFor(() => expect(calls).toContain("GET /api/health"));
    await new Promise((r) => setTimeout(r));
    await stop();
    release([]);
    await first;
    await solve(excelWith());
    expect(calls).toContain("POST /api/run");
  });
});

describe("another solve is already running", () => {
  const busy = () => json({ busy: true, run_id: "old" }, 409);

  it("offers to stop it", async () => {
    route({ "/api/run": busy, "/api/run/old/cancel": () => json({ ok: true }) });
    answerConfirm(true);
    await solve(excelWith());
    expect(calls).toContain("POST /api/run/old/cancel");
    expect(runStore.get().phase).toBe("stopped");
  });

  it("leaves it alone when declined", async () => {
    route({ "/api/run": busy });
    answerConfirm(false);
    await solve(excelWith());
    expect(calls).not.toContain("POST /api/run/old/cancel");
    expect(runStore.get().error?.message).toContain("still running");
  });
});

it("shows the helper's own message when a run can't start", async () => {
  route({ "/api/run": () => json({ error: "Could not start the model. (no python3)" }, 500) });
  await solve(excelWith());
  expect(runStore.get().error?.message).toBe("Could not start the model. (no python3)");
});

it("says results may be partly written when a write fails", async () => {
  itemsStore.set({ [SHEET]: [item("out", "A1")] });
  route({
    "/api/run": () => json({ run_id: "r" }),
    "/api/run/r/events": () => stream(sse(result([{ name: "out", value: [5] }]))),
  });
  const excel: ExcelAdapter = {
    ...excelWith(),
    writeRanges: async () => {
      throw new Error("Excel said no");
    },
  };
  await solve(excel);
  const { error } = runStore.get();
  expect(error?.message).toContain("some of them may already be in the sheet");
  expect(error?.message).toContain("Excel said no");
});

it("keeps a read-only item's cells when a result lands on them", async () => {
  const mock = createMockAdapter({ A1: 15, B1: 16 });
  itemsStore.set({
    [SHEET]: [{ ...item("Kar", "A1:B1"), readOnly: true }, item("Uretim", "A1:B1")],
  });
  route({
    "/api/run": () => json({ run_id: "r" }),
    "/api/run/r/events": () => stream(sse(result([{ name: "Uretim", value: [200, 100] }]))),
  });
  await solve(mock.adapter);
  expect(mock.cells.get("A1")).toBe(15);
  expect(mock.cells.get("B1")).toBe(16);
  expect(runStore.get().warnings?.join(" ")).toContain("read-only Kar");
});

it("shows every log line without re-rendering for each chunk", async () => {
  const lines = Array.from({ length: 2000 }, (_, i) => ({ type: "log", line: `line ${i}` }));
  route({
    "/api/run": () => json({ run_id: "r" }),
    "/api/run/r/events": () => stream(...lines.map((l) => sse(l)), sse(result([]))),
  });
  let updates = 0;
  const unsubscribe = runStore.subscribe(() => updates++);
  await solve(excelWith());
  unsubscribe();
  expect(runStore.get().phase).toBe("finished");
  expect(runStore.get().log).toHaveLength(2000);
  expect(updates).toBeLessThan(50);
});

it("shows a lone line soon, even if nothing follows it", async () => {
  let finish = () => {};
  route({
    "/api/run": () => json({ run_id: "r" }),
    "/api/run/r/events": () =>
      new Response(
        new ReadableStream({
          start(controller) {
            controller.enqueue(encoder.encode(sse({ type: "log", line: "building" })));
            finish = () => {
              controller.enqueue(encoder.encode(sse(result([]))));
              controller.close();
            };
          },
        }),
      ),
  });
  const solving = solve(excelWith());
  await vi.waitFor(() => expect(runStore.get().log).toContain("building"), { timeout: 1000 });
  finish();
  await solving;
});

function item(name: string, cellRange: string) {
  return { name, type: "list" as const, cellRange, indexRanges: [], readOnly: false };
}
function result(changed: unknown[]) {
  return { type: "result", elapsed: 1, status: "Optimal", objective: 1, changed, warnings: [] };
}
