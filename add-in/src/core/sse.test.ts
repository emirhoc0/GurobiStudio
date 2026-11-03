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
