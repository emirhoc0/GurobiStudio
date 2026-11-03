# GurobiStudio

Build and solve Gurobi optimization models from Excel for Mac. Bind sheet cells to Python names, write
a `gurobipy` model in the task pane, and solve. Results are written back to the sheet. Models written
for SolverStudio mostly run unchanged, apart from Python 2 syntax.

It has two parts: an Office add-in, and a small Python helper that runs on your Mac and only listens
on `127.0.0.1`.

## Requirements

- Excel for Mac (tested on 16.112)
- Python 3.13 and Node 22
- A Gurobi license. Without one, `gurobipy` falls back to its size-limited license.

## Install

```sh
scripts/install.sh
scripts/run-helper.sh
```

Then restart Excel and choose Home > Add-ins > GurobiStudio. The buttons appear on the Data tab.

Install is safe to re-run. The first run may ask for your password to trust a local certificate.

| Option | Effect |
|---|---|
| `GS_PORT=9000` | Use another port. Restart the helper and Excel afterwards. |
| `GS_PLACEMENT=custom` or `home` | Put the buttons on their own tab or on Home instead of Data. |
| `GS_LAUNCH_AGENT=1` | Start the helper at login and keep it running. |

To remove it, run `scripts/uninstall.sh`. Your workbooks and the data folder are kept.

## Try an example

Choose More > Examples on the ribbon, or File > View Examples on the Model tab, and insert one. Each
adds a new sheet with data, data items and a model. Then click Solve Model.

## Write your own

1. On **Data Items**, add a name for each range. Types are Single value, List, Indexed and Table.
2. On **Model**, write the model. Data items are ordinary Python variables:

   ```python
   from SolverStudio import *
   from gurobipy import *

   m = Model()
   x = m.addVars(supply.keys(), demand.keys(), obj=cost)
   m.addConstrs(x.sum(s, "*") <= supply[s] for s in supply)
   m.addConstrs(x.sum("*", d) == demand[d] for d in demand)
   m.optimize()
   for key in x:
       flow[key] = x[key].X
   ```

3. Click **Solve Model**, or press ⌘↵. ⌘. stops a run.

Any item the model changes is written back to its cells and marked with `*` in the log. Lock an item
as read-only to protect its cells.

## Good to know

- **Models run as Python with your full access.** New or changed code shows a prompt first, so read
  code from a workbook you didn't write before running it. Nothing runs when a workbook opens.
- **Save the workbook to keep your work.** Models, items and solver settings are stored inside it.
- **Formulas are protected.** If a result would replace a formula, you're asked first.
- **One model per worksheet.** A copied sheet starts empty. Solver settings apply to the whole workbook.
- Run logs and files are in `~/Library/Application Support/GurobiStudio/runs/`, or use View Working
  Files. The newest 10 runs are kept, which you can change in Settings.

## Big models

Tested up to 1M input cells and 100k written cells on a test sheet. In real Excel it has only been
proven at 10k cells so far, see `docs/TESTING.md`. To keep large models quick:

- Use `addVars` and `quicksum`, not repeated `+=` on an expression.
- Don't print a line per row. The pane keeps the last 5,000 lines, and the full log is in the run folder.
- Use a bounded range, a Table column or a named range. Ranges over 5M cells are refused.
- Mark inputs read-only. A result can then never replace them, even if another item's range is wrong.

## Troubleshooting

| Problem | Try |
|---|---|
| Red "Helper unreachable" badge | Run `scripts/run-helper.sh`. |
| Solve reports a license problem | Settings > Check License. A WLS license needs internet access. |
| "Size-limited" license | Gurobi didn't find your license. Set its path in Settings. |
| Buttons are missing | Home > Add-ins > GurobiStudio. |
| Pane is stale after an update | See "Reload after a manifest change" in `docs/TESTING.md`. |

## Develop

See `CLAUDE.md` for commands and architecture, `docs/DECISIONS.md` for why things are the way they
are, and `docs/TESTING.md` for the checks in real Excel. `npm run e2e` in `add-in` runs the browser
acceptance test against the real helper and Gurobi.
