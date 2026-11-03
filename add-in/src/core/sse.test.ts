import { expect, it } from "vitest";
import { createSseParser } from "./sse";

it("reassembles events split across chunks, even inside a character", () => {
  const seen: unknown[] = [];
  const feed = createSseParser((e) => seen.push(e));
  const bytes = new TextEncoder().encode('data: {"a":"ç"}\n\ndata: {"b":2}\n\n');
  for (const byte of bytes) feed(new Uint8Array([byte]));
  expect(seen).toEqual([{ a: "ç" }, { b: 2 }]);
});

it("handles several events in one chunk", () => {
  const seen: unknown[] = [];
  createSseParser((e) => seen.push(e))(new TextEncoder().encode("data: 1\n\ndata: 2\n\n"));
  expect(seen).toEqual([1, 2]);
});

it("delivers an event whose chunk ends exactly on its blank line", () => {
  const seen: unknown[] = [];
  createSseParser((e) => seen.push(e))(new TextEncoder().encode("data: 1\n\n"));
  expect(seen).toEqual([1]);
});

it("finds a blank line split across two chunks", () => {
  const seen: unknown[] = [];
  const feed = createSseParser((e) => seen.push(e));
  feed(new TextEncoder().encode('data: "a\\nb"\n'));
  expect(seen).toEqual([]);
  feed(new TextEncoder().encode('\ndata: "c"\n\n'));
  expect(seen).toEqual(["a\nb", "c"]);
});

it("reassembles a large event in about linear time", () => {
  const seen: { x: string }[] = [];
  const feed = createSseParser<{ x: string }>((e) => seen.push(e));
  const bytes = new TextEncoder().encode(`data: ${JSON.stringify({ x: "a".repeat(100e6) })}\n\n`);
  const started = performance.now();
  for (let i = 0; i < bytes.length; i += 65536) feed(bytes.subarray(i, i + 65536));
  expect(seen[0].x.length).toBe(100e6);
  expect(performance.now() - started).toBeLessThan(3000); // about 9 s when it was quadratic
});
