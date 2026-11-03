# Decisions

Each entry says what was chosen and why. "Proven" means tested in real Excel for Mac.

## Differences from the brief

- **Python 2 models need editing.** SolverStudio's published Gurobi example uses `print 'x'`, a syntax
  error in Python 3. The error card names the line. Everything else in that example works.
- **Office.js loads from Microsoft's CDN**, as Microsoft's docs direct, so "no network except WLS"
  has this one exception. The add-in makes no other external calls and has no telemetry.
- **One model per worksheet**, keyed by `worksheet.id`. Renames and moves are safe; a copied sheet
  starts empty. Settings are per workbook.
- **API additions:** `POST /api/run/{id}/reveal`, `approved` plus a 409 on `POST /api/run`, and
  `reserved_names` on `/api/health`.
- **Streaming uses `fetch`** because `EventSource` cannot send an `Authorization` header.
- **Button sizes are Office's choice.** An add-in manifest cannot mark a ribbon button large or
  small. Excel happened to render Stop small at the window width tested.

## From section 4 of the brief

- **Data direction:** every item is read before a run; changed items are written back. Optional
  per-item read-only lock.
- **Log layout and solver options UI:** designed fresh, see the Output and Settings tabs.

## Infrastructure

- **Certificates:** `office-addin-dev-certs` with `--days 800`, which applies to both the CA and the
  server certificate (read from the package source). macOS rejects TLS certificates valid for over
  825 days. The CA goes in the login keychain and is removed by `scripts/uninstall.sh`. Its usage
  reporting defaults to off and raises no prompt.
- **Ribbon commands use a shared runtime.** Microsoft's docs state that no `Action` may have a
  `TaskpaneId` in this setup, and that `showAsTaskpane` still works for sideloaded add-ins.
- **Helper data folder:** `~/Library/Application Support/GurobiStudio/`.
- **Port:** 8765 by default. It is written into the manifest, so changing it means re-running install.
- **TypeScript 6.0.3**, not 7, because typescript-eslint requires a version below 6.1.

## Security

- **The API token needs a page key.** The brief injects the token into the served HTML, but any local
  program could then fetch the page and read it. The page now requires `?k=<page key>`. The key lives
  only in a 0600 file in the helper data folder and in the private sideloaded manifest. Without it
  the page returns 403 and never contains the token.
- **Not covered:** a program running as you outside the sandbox can read those files anyway. It can
  already run anything as you, so this adds no new risk. No peer-UID check on loopback.

## Proven by M1 diagnostics (Excel 16.112.4 for Mac)

- **Persistence:** a custom XML part round-trips tabs, CRLF, non-ASCII and `]]>` exactly within a
  session. Save-and-reopen is still to be proven at M5. A very hidden sheet also works as a fallback,
  but Excel refuses to delete one until it is made visible.
- **Excel Tables:** a name defined as `=Table[Col]` has type `Range`, and both it and the table API
  grow with the table (`B2:B3` to `B2:B4`). Acceptance test 7 should work through a plain named range.
- **Origin header:** `https://localhost:8765`, so the strict check stands.
- **Write semantics:** `null` keeps a cell, including a formula; `""` clears it.
- **Cell values:** `#N/A` reads as type `Error`; `3` serializes as `3`; booleans stay booleans.
- **Streaming:** `fetch` delivers SSE incrementally (chunks at 11, 411, 822 and 1224 ms).
- **Dialog:** two-way messaging works, so the pop-out editor can use `messageChild`.
- **Theme:** `Office.context.officeTheme` is available, with `isDarkTheme`.
- **Worksheet id:** survives a rename.

- **Ribbon:** the group on the built-in Data tab (`TabData`) works, so no fallback was needed. Stop
  renders small and the other three large. Excel for Mac showed no group titles in the test
  screenshot, so the name "SolverStudio" may not be visible on the ribbon.

## Data items (M3)

- **Trailing blanks and the index.** The index decides how many entries there are. A cell range longer
  than the index is fine if the extra cells are blank; values past the end are an error. Every key
  keeps its cell even when blank, so an empty output column still writes back.
- **Multi-column lists** drop trailing blank rows only. A blank column stays as `None` in each tuple,
  so the tuple length doesn't change once results arrive.
- **Blanks inside a list** are `None`, so positions hold. Blanks in dicts are omitted, as the brief says.
- **A blank or repeated index key is an error**, since either would silently lose data.
- **Whole rows and columns (`A:A`) are refused.** Reading a million cells freezes Excel. Use a bounded
  range, a Table column or a named range.
- **Overlapping ranges warn, not block.** Only writable items count. Checked on resolved addresses, so
  named ranges and tables are covered.
- **`Application` is refused as an item name**, because the runner puts the shim object there.
- **Result strings are stored as exact text.** Every string is written with a leading apostrophe, so
  `=1+1` can't run as a formula and `007` stays `007`. Excel keeps the apostrophe as a hidden "text"
  marker, visible only in the formula bar. Models can't write formulas; add an explicit opt-in if
  that's ever wanted. A list of risky prefixes was tried first and dropped after security review.
- **Named range creation** names the Cell Range only, at workbook scope, using the item's name. A name
  Excel refuses (`x1`, `c`, `r`) shows Excel's message.
- **Formula names.** Where Excel gives a name's values but no address, the item still reads, and a
  change is reported as a warning instead of written. Whether `OFFSET` names land here is checked by
  diagnostics.
- **The pane stays on its sheet while a form has unsaved edits**, so a range can be picked from another
  sheet, and after saving until the next sheet click, so the saved item doesn't seem to vanish. Solve
  uses the active sheet, as ever.

## Proven by M3 diagnostics (Excel for Mac, 2026-11-03)

- **Reading and writing through the adapter:** addresses, quoted sheet names, `Table[Col]`,
  `Table[[A]:[B]]`, a bare table name, named ranges, error cells and the selection address all work.
  A bad reference is reported alone and the others still read. `null` keeps a formula and `""` clears.
- **Result strings stay exact text:** `=1+1`, ` =1+1`, the full-width `＝1+1`, `+1`, `007`, `1/2` and
  `'x` all read back unchanged with type `String`.
- **`OFFSET` names work fully.** One made in Name Manager has type `Range` and resolves to an address,
  so it can be read and written like any other range, with no fallback. The API still can't create one:
  `names.add` rejected it with `InvalidArgument`, so make those in Name Manager. This Excel shows the
  formula with `;` separators, which may be why, but that wasn't tested.
- **Not exercised in Excel:** the values-only path for names with no address. No test name reached it.

## Editor and prompts (M4)

- **Data item names are chips above the editor**, not a side list, because the pane is too narrow.
- **The trust gate is advisory.** Anything holding the token can send `approved`. That program already
  runs as you, so this adds no new risk. The gate stops a workbook from running new code unseen.
- **Formula prompt is one dialog per solve.** "Keep formulas" still writes every other result.
- **The pane keeps the last 5,000 log lines.** The full log stays in the run folder.

## Persistence and settings (M5)

- **State is one custom XML part** (`urn:gurobistudio:state`) holding JSON with `schema: 1`.
- **Saves are debounced 300 ms and never overlap.** A failure shows one toast and retries every 2 s.
- **Unreadable or newer state turns saving off** and shows an error, so it's never overwritten.
- **Nothing saves until the load finishes**, and `solve()` waits for it. Nothing runs on open.
- **Sheet ids:** if no stored id matches a current sheet, entries are matched by name. If any id
  matches, none are remapped, so a new sheet can't inherit a deleted sheet's model.
- **The port is read-only in Settings**, because changing it means rewriting the manifest.
  `install.sh` writes a `port` file that the helper reads, so the two can't disagree.
- **Check License returns derived facts only**, never raw log lines, and refuses during a solve
  because a solve may hold the only WLS session.
