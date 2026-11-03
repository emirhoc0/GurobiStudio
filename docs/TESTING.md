# Manual checks in Excel for Mac

Automated tests can't click inside Excel. A box stays unticked until someone has run the check in real
Excel. Ticked boxes dated 2026-11-03 were driven by AppleScript and `cliclick` on Excel 16.112.4.

## Reload after a manifest change

1. Quit Excel.
2. Run `scripts/install.sh`.
3. Delete the contents of `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`, if it exists.
   Otherwise delete `~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/`.
4. Re-run `scripts/install.sh`, since clearing can also remove the sideloaded manifest.
5. Open Excel, then Home > Add-ins > GurobiStudio.

To inspect the task pane: `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`,
then right-click in the pane and choose Inspect Element.

## M1 pipeline

The manifest URL carries the page key, so re-run install after deleting the helper data folder.

- [x] The Data tab shows Edit Data, Show Model, Solve Model, Stop, More (2026-11-03).
- [ ] Edit Data opens the pane on Data Items.
- [x] Show Model opens Model, on the second click after a reopen (see DECISIONS, 2026-11-03).
- [x] The pane shows a green Gurobi 13.0.3 badge (2026-11-03).
- [ ] Settings > Run diagnostics finishes, and "temporary sheet and names removed" passes.

## M3 data items and solve

Settings > Run diagnostics now also tests the real cell, name and selection paths. To include the
`OFFSET` check, first make a name in Formulas > Name Manager called `gs_offset`, referring to
`=OFFSET($A$1,0,0,3,1)`. The report goes to `diag/latest.json` in the data folder.

- [x] Diagnostics: every "adapter" row and "OFFSET name made in Name Manager" pass (2026-11-03).

On a blank sheet, enter this:

| | A | B | C |
|---|---|---|---|
| 1 | | c1 | c2 |
| 2 | w1 | 1 | 4 |
| 3 | w2 | 3 | 2 |
| 5 | w1 | 20 | |
| 6 | w2 | 30 | |
| 8 | c1 | 25 | |
| 9 | c2 | 15 | |

Then in Data Items add:

| Name | Type | Index | Cell Range |
|---|---|---|---|
| `cost` | Table | row `A2:A3`, column `B1:C1` | `B2:C3` |
| `supply` | Indexed | `A5:A6` | `B5:B6` |
| `demand` | Indexed | `A8:A9` | `B8:B9` |
| `flow` | Table | row `A2:A3`, column `B1:C1` | `E2:F3` |
| `status` | Single value | | `H1` |

Paste this on the Model tab and click Solve Model:

```python
from SolverStudio import *
from gurobipy import *

m = Model("transport")
x = m.addVars(supply.keys(), demand.keys(), obj=cost)
m.addConstrs(x.sum(w, "*") <= supply[w] for w in supply)
m.addConstrs(x.sum("*", c) == demand[c] for c in demand)
m.optimize()
for k in x:
    flow[k] = x[k].x
status = "Solved"
```

- [ ] Output shows Optimal and Objective 65, and `flow*, status*` as changed.
- [ ] E2:F3 holds 20, 0, 5, 15, with no formatting changed. H1 says Solved.
- [ ] Put a formula in `B5` (`=10+10`) and solve again: the formula is still there.
- [ ] Pick from selection fills a range field from the sheet.
- [ ] Click another sheet with a half-filled form and pick there: the field gets `Sheet2!A1:A3`, and
      the pane stays on the first sheet. After Add, the next sheet click shows that sheet's own items.
- [ ] Create named range on an item with an address: the name appears in Name Manager and the field
      switches to it. A name like `x1` shows Excel's refusal.
- [ ] Put `=NA()` in B2: Solve stops and names the cell, and nothing is written.

Excel Table (acceptance test 7):

- [x] A List item on `Demand[Qty]` previews `[1, 2]`. Typing a row under the Table and clicking Refresh
      gives `[1, 2, 3]` (2026-11-03).
- [ ] Same with an Indexed item using `Demand[Customer]` as its index.
- [ ] Bind a named range to a table column, add a row, and check again.

## M4 to M6

Examples:

- [x] More > Examples opens the picker. Insert Transportation adds a formatted sheet with its model and
      items (2026-11-03).
- [x] Solve shows the trust prompt, then Optimal. Cells read back by AppleScript equal the brute-force
      answer: 20, 0, 10 / 0, 30, 15, total 365 (2026-11-03).
- [ ] Diet, Knapsack and Assignment in Excel. All four pass in `npm run e2e` on the mock sheet.

Prompts and running:

- [x] Changed code prompts again. Unchanged code doesn't (2026-11-03).
- [x] A formula that would change shows one prompt naming the cell. Keep leaves it, Overwrite replaces it.
      A formula equal to the result shows nothing (2026-11-03).
- [x] Ribbon Stop during a long MIP: chip says Stopped, Gurobi logs the interrupt, runner exits at once,
      no cells written (2026-11-03).
- [x] Pop out opens the model. Text typed there reaches the pane (2026-11-03).
- [ ] Pop out with a very large model, for the `messageParent` size limit.
- [ ] ⌘↵ and ⌘. inside Excel, and error-card jump to line. All pass in `npm run e2e`.

Saving:

- [x] Save, close, reopen: model and items restored, nothing runs, approval kept (2026-11-03).
- [x] Solve Model as the first click after a reopen opens the pane and runs (2026-11-03).
- [ ] Tabs, CRLF and non-ASCII through a real reopen. Passes in `npm run e2e` and in M1's session test.
- [ ] Settings tab in Excel: parameters, Check License, helper paths. All pass in `npm run e2e`.

Light theme and install:

- [x] The pane follows Excel's dark theme (2026-11-03).
- [ ] Excel in light mode, and switching while the pane is open.
- [x] `scripts/install.sh` twice: exit 0, with the same manifest, page key, token and certificate, and
      no certificate prompt (2026-11-03).
- [ ] `GS_LAUNCH_AGENT=1 scripts/install.sh`: the helper runs after login, and `scripts/uninstall.sh`
      removes it. Never run. Stop any running helper first, because they share a port. Then check
      `helper.log` in the data folder, since macOS may block a launchd job from reading `~/Desktop`.
