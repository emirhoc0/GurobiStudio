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
  small, so Stop may not render small. Excel sizes controls to fit the ribbon.

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

## To be settled by M1 diagnostics

- Ribbon placement (`TabData`, then custom tab, then Home).
- Persistence: custom XML part or very hidden sheet.
- `Table1[Col]` and `OFFSET`-style named ranges.
- The `Origin` header Excel's web view really sends.
