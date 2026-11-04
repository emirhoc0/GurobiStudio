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
- **Overlapping ranges warn, not block.** A pair counts if at least one item is writable. Checked on
  resolved addresses, so named ranges and tables are covered.
- **A read-only item's cells are never replaced by a result**, even when another item points at them.
  The other cells are still written, with a warning. A real workbook lost its profit row this way: a
  writable item was bound to the same cells as the read-only input, and no warning appeared.
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

## Examples and install (M6)

- **Examples are JSON bundled into the add-in**, so they need no extra API. All data is invented for
  this project. Their `expected` answers come from brute force, not Gurobi, and each optimum was
  checked to be unique, so the acceptance test can't fail on an alternate optimum.
- **Inserting an example sets its model before activating the sheet**, so the editor never shows the
  template first.
- **`npm ci` is retried once in install.** It failed once with `ENOTEMPTY` deleting `node_modules`, and
  the retry passed. What touched the folder wasn't identified.
- **The login item is on by default**, replacing opt-in, because the helper serves the pane and the
  ribbon buttons. With it off they just fail to load, with no badge to explain why. `GS_LAUNCH_AGENT=0`
  opts out. A repo under `~/Desktop`, `~/Documents` or `~/Downloads` stays manual unless
  `GS_LAUNCH_AGENT=1`, since macOS may limit launchd jobs there. That limit wasn't tested.
- **Known issue: the first Show Model click after reopening a workbook does nothing.** A second click
  opens the pane. Seen every time in real Excel, 5 of 5 tries. Solve Model as the first click worked
  then, but on two freshly opened workbooks on 2026-11-03 it showed no pane within 12 s, and a second
  click did. Nothing is at risk, because a solve can't start twice. Ruled out: the action being associated too late (it's now at the top
  level), `showAsTaskpane` running before `onReady`, and whether the pane was open last session. Edit
  Data, Settings and Examples share this code path, but only Show Model was tried.
- **Gurobi reports "no expiry" as `0`, `99991231` or `99999999`.** Check License shows no date for these.

## Proven in real Excel (M4 to M6, Excel 16.112.4, 2026-11-03)

- **Save, close and reopen:** the model and items come back from the file, nothing runs on open, and a
  model approved before the reopen runs without a new prompt, since the workbook id is stored.
- **The restored model lands on the right sheet.** Whether through a surviving worksheet id or the remap
  by name wasn't checked.
- **The page key is not stored in a saved workbook.** Searched the unzipped file.
- **The pop-out's short edits reach the pane.** The size limit of `messageParent` was not tested in real
  Excel. The 300 KB test only covers the browser mock.

## Going public

- **Approval covers text-valued Gurobi parameters.** They come from the workbook, and some name a file
  or a server (`LogFile`, `ComputeServer`), so they join the fingerprint and appear in the prompt.
  Numbers can't do that, so changing `TimeLimit` doesn't prompt again. Chosen over a list of dangerous
  names, which would go stale with new Gurobi versions.
- **`install.sh` checks certificate trust itself.** `office-addin-dev-certs install` exits 0 when it
  fails to install the CA, which would report success and leave a pane that can't load.
- **The pane warns when it isn't in Excel.** If `office.js` can't load, it falls back to the mock
  adapter, which browser tests rely on. Inside Excel that looked normal while nothing was saved.
- **The README recommends an academic WLS license.** A named-user license needs `grbgetkey`, which the
  pip package doesn't include.

## Big data

Measured numbers and what is still unproven are in `docs/TESTING.md`.

- **Target:** 1M input cells and 100k written cells. JSON stays the wire format, since 1M cells cost
  about 30 MB and 0.3 s to encode.
- **Office.js reads and writes go in row chunks of 100k cells or fewer per sync.** An estimate, not a
  measured limit: real Excel has only run ranges too small to split.
- **Ranges over 5M cells are refused before their values load.** Also an estimate, chosen to stop
  `A1:Z1048576`, which gets past the whole-column refusal.
- **Calculation is suspended on every sync but the last**, so a formula over the results recalculates
  once. Whether dependents end up correct across chunks still needs the 100k-written check.
- **A chunked write can stop midway**, so that error says some results may already be in the sheet.
- **Run events are consumed as they're sent, and at most 50,000 wait in memory.** A very chatty model
  loses the oldest unsent lines from the pane, never the result, and `log.txt` keeps every line. There
  is one reader by design, so reopening a run's stream returns nothing.
- **The snapshot is a one-level copy.** Item values are scalars and tuples, so nothing deeper can change.
- **A re-solve can report an unchanged item as changed.** Excel keeps 15 significant digits, so a value
  read back can differ slightly from Gurobi's. The only effect is a harmless rewrite and a `*`.
- **Another solve still running now offers to stop it**, since a reloaded pane has lost its run id.
- **The Data Items tab still reads whole ranges to preview them.** Left alone until real Excel shows
  whether that's noticeable at 1M cells.
