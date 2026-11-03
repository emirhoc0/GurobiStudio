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

## Still to settle

- `OFFSET`-style names. `names.add` rejects the formula even with a quoted sheet name, so the API
  can't create one. Reading one made in Excel's Name Manager is tested at M3.
