# GurobiStudio

Build and solve Gurobi optimization models from Excel on a Mac, like SolverStudio does on Windows. Bind
sheet cells to Python names, write a `gurobipy` model in a pane inside Excel, click Solve, and the
results are written back to your sheet.

Not affiliated with or endorsed by Gurobi Optimization or Microsoft.

## What you need

- A Mac, and Excel for Mac (tested on 16.112). Open Excel at least once before installing.
- Python 3.13 and Node 22.12 or newer. Step 1 below installs them.
- Your own Gurobi license. Step 2 below explains how, and it's free for students.
- Internet access while using it. The pane loads Microsoft's Office library from the web, and an
  academic WLS license checks in online.

## Setup

Takes about 10 minutes. Run the commands in the Terminal app.

### 1. Install Python and Node

With [Homebrew](https://brew.sh):

```sh
brew install python@3.13 node
```

Or use the installers from [python.org](https://www.python.org/downloads/) (choose a 3.13 version) and
[nodejs.org](https://nodejs.org).

### 2. Get your own Gurobi license

Everyone needs their own. Don't copy someone else's file: `gurobi.lic` holds private credentials, so
never share it or commit it to a repo.

1. [Register at Gurobi](https://portal.gurobi.com/iam/register/) with your university email, as an
   academic.
2. [Request an academic license](https://portal.gurobi.com/iam/licenses/request?type=academic) and
   choose the **WLS** (Web License Service) one. Do this on the campus network or VPN, because Gurobi
   checks that you're connecting from a university.
3. Go to [license.gurobi.com](https://license.gurobi.com), find the license and click **Download** to
   get `gurobi.lic`.
4. Save it in your home folder:

   ```sh
   mv ~/Downloads/gurobi.lic ~/gurobi.lic
   ```

No license yet? You can skip this step. Gurobi's built-in license works for models up to 2,000 variables
and 2,000 constraints, and you can add your license at any time without reinstalling.

A named-user academic license also works, but getting one needs the full Gurobi installer for its
`grbgetkey` tool. WLS is simpler.

### 3. Download and install GurobiStudio

```sh
git clone https://github.com/emirhoc0/GurobiStudio.git ~/GurobiStudio
cd ~/GurobiStudio
scripts/install.sh
```

Keep it in your home folder as shown. macOS may block the automatic start from Desktop, Documents and
Downloads.

macOS may ask two things during the install. Accept both:

- **Your password**, to trust a certificate created on your Mac. Excel only loads add-ins over HTTPS.
- **Terminal would like to access data from other apps**, to put the add-in in Excel's folder.

If something's missing, the installer stops and tells you what to run. It's safe to run again.

### 4. Open it in Excel

1. Quit Excel completely (⌘Q) and open it again.
2. Choose **Home > Add-ins > GurobiStudio**. The buttons appear on the **Data** tab.
3. In the pane, open **Settings** and click **Check License**. It should report your license. If it says
   size-limited, Gurobi didn't find `~/gurobi.lic`.

### 5. Solve an example

Choose **More > Examples** on the ribbon and insert one. It adds a new sheet with data and a model.
Click **Solve Model**, read the code when asked, and click **Run**. The results appear on the sheet.

## Write your own model

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

**A result only reaches the sheet if your model assigns it to a data item**, like `flow` above. Any item
the model changes is written back to its cells and marked with `*` in the log. Lock an item as read-only
to protect its cells.

Models written for SolverStudio mostly run unchanged, apart from Python 2 syntax such as `print 'x'`.

Good to know:

- **Save the workbook to keep your work.** Models, items and solver settings are stored inside it.
- **Formulas are protected.** If a result would replace a formula, you're asked first.
- **One model per worksheet.** A copied sheet starts empty. Solver settings apply to the whole workbook.
- Run logs and files are in `~/Library/Application Support/GurobiStudio/runs/`, or use View Working
  Files. The newest 10 runs are kept, which you can change in Settings.

## Sharing workbooks

Send the `.xlsx` as usual. The model and data items travel inside it, and your license doesn't. The other
person needs GurobiStudio and their own license.

**A model runs as Python with full access to your Mac, so only run code you've read.** Nothing runs when a
workbook opens. New or changed code shows a prompt first, along with any file or server settings the
workbook passes to Gurobi. Be careful with a workbook from someone you don't know.

## The helper

GurobiStudio has two parts. The **add-in** is the pane inside Excel. The **helper** is a small Python
server that runs your models with Gurobi, since Excel can't run Python itself. It only listens on
`127.0.0.1`, so nothing outside your Mac can reach it.

The helper also serves the pane and the ribbon buttons, so neither works when it's off. The installer sets
it to start at login. It uses almost nothing while idle. Closing Excel doesn't stop it.

```sh
# Off, until you turn it on again or next log in
launchctl bootout gui/$(id -u)/com.gurobistudio.helper

# On again
launchctl bootstrap gui/$(id -u) ~/Library/LaunchAgents/com.gurobistudio.helper.plist
```

To be sure a model stops cleanly, stop it first with ⌘. in the pane. That lets Gurobi release its license.

**Prefer no background process?** Install with `GS_LAUNCH_AGENT=0 scripts/install.sh`. Then run
`scripts/run-helper.sh` before opening Excel and leave that Terminal window open. Ctrl+C stops it.

### Install options

| Option | Effect |
|---|---|
| `GS_PORT=9000` | Use another port. Restart Excel afterwards. |
| `GS_PLACEMENT=custom` or `home` | Put the buttons on their own tab or on Home instead of Data. |
| `GS_LAUNCH_AGENT=0` | Don't start the helper at login. |
| `GS_LAUNCH_AGENT=1` | Start it at login even from Desktop, Documents or Downloads. |

## Update

```sh
cd ~/GurobiStudio
git pull
scripts/install.sh
```

Then restart Excel. If the pane still looks old, see "Reload after a manifest change" in
[docs/TESTING.md](docs/TESTING.md).

## Uninstall

```sh
cd ~/GurobiStudio
scripts/uninstall.sh
```

This removes the add-in, the login item and the trusted certificate. To remove everything else, delete
`~/GurobiStudio` and `~/Library/Application Support/GurobiStudio`. Your workbooks and `~/gurobi.lic` are
never touched.

## Troubleshooting

| Problem | Try |
|---|---|
| The pane won't load, or the buttons do nothing | The helper is probably off. Run `pgrep -fl 'm gurobistudio$'`. If nothing prints, run `scripts/install.sh` again, or `scripts/run-helper.sh` if you opted out of the login item. |
| Red "Helper unreachable" badge | The helper stopped after the pane opened. Same fix, then close and reopen the pane. |
| First button click after opening a workbook does nothing | Known issue. Click again. |
| Buttons are missing | Home > Add-ins > GurobiStudio. |
| "Not connected to Excel" in the pane | The pane couldn't load Microsoft's Office library. Check your internet connection, then reopen the pane. Anything typed there isn't saved. |
| Model too large for size-limited license | Gurobi didn't find your license. Check that the file is at `~/gurobi.lic`, or set its path in Settings. Then use Check License. |
| Solve reports a license problem | Settings > Check License. A WLS license needs internet access. It also expires, so renew it on Gurobi's portal and download a new file. |
| Academic license request is refused | Connect through the campus network or VPN and try again. |
| "Another program is using port 8765" | `GS_PORT=9000 scripts/install.sh`, then restart Excel. |
| Add-in stops loading after about two years | The local certificate lasts 800 days. Run `scripts/install.sh` again. |
| Pane is stale after an update | See "Reload after a manifest change" in [docs/TESTING.md](docs/TESTING.md). |
| Anything else with the helper | Read `~/Library/Application Support/GurobiStudio/helper.log`. |

## Big models

Tested up to 1M input cells and 100k written cells on a simulated sheet. In real Excel it has only been
proven at 10k cells so far, see [docs/TESTING.md](docs/TESTING.md). To keep large models quick:

- Use `addVars` and `quicksum`, not repeated `+=` on an expression.
- Don't print a line per row. The pane keeps the last 5,000 lines, and the full log is in the run folder.
- Use a bounded range, a Table column or a named range. Ranges over 5M cells are refused.
- Mark inputs read-only. A result can then never replace them, even if another item's range is wrong.

## How a solve works

1. The add-in reads the cells bound to your data items and sends them with the model to the helper.
2. The helper runs the model in a separate Python process, where each data item is a normal variable.
3. The log streams into the pane while Gurobi runs.
4. When the model ends, the helper compares the data items to their starting values. The add-in writes
   back only the ones that changed.

## For developers

- [docs/TESTING.md](docs/TESTING.md): what has been checked in real Excel
- [docs/DECISIONS.md](docs/DECISIONS.md): why things are the way they are
- [CLAUDE.md](CLAUDE.md): commands and architecture
