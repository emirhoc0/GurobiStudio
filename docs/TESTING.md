# Manual checks in Excel for Mac

Automated tests can't click inside Excel. These stay **unverified** until someone runs them.

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

- [ ] The Data tab shows a SolverStudio group with Edit Data, Show Model, Solve Model, Stop, More.
- [ ] Edit Data opens the pane on Data Items. Show Model opens Model.
- [ ] The pane shows a green Gurobi 13.0.3 badge.
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

- [ ] Make A8:B9 a Table with headers on row 7 (name it `Demand`, columns `Customer`, `Qty`). Set the
      `demand` item to index `Demand[Customer]`, cells `Demand[Qty]`. Add a row and check the preview
      shows the new entry. Then bind a named range to a table column, add a row, and check again.
