/** Incremental SSE parser. Chunks can split anywhere, including inside a multi-byte character. */
export function createSseParser<T>(onEvent: (event: T) => void) {
  const decoder = new TextDecoder();
  // Parts, not one growing string, so a result that spans thousands of chunks stays linear.
  let parts: string[] = [];
  let carry = ""; // a trailing newline, which may be half of the next chunk's blank line
  return (chunk: Uint8Array) => {
    const text = carry + decoder.decode(chunk, { stream: true });
    carry = "";
    let start = 0;
    let end: number;
    while ((end = text.indexOf("\n\n", start)) !== -1) {
      parts.push(text.slice(start, end));
      const block = parts.join("");
      parts = [];
      start = end + 2;
      if (block.startsWith("data: ")) onEvent(JSON.parse(block.slice(6)) as T);
    }
    let rest = text.slice(start);
    if (rest.endsWith("\n")) {
      carry = "\n";
      rest = rest.slice(0, -1);
    }
    if (rest) parts.push(rest);
  };
}
