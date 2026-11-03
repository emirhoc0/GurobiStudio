# GurobiStudio

A SolverStudio alternative for Excel on Mac. Bind sheet cells to Python names, write a `gurobipy` model
in the Excel task pane, and solve. Results are written back to the sheet.

First time here? [docs/README.md](docs/README.md) covers install, requirements and writing models.

## How it works

There are two parts:

- **The add-in** is the pane inside Excel. It reads and writes cells and shows the log.
- **The helper** is a small Python server on your Mac. It runs your models with Gurobi. It only listens
  on `127.0.0.1`, so nothing outside your Mac can reach it.

Excel can't run Python itself, so the add-in needs the helper to solve. When you click **Solve Model**:

1. The add-in reads the cells bound to your data items and sends them with the model to the helper.
2. The helper runs the model in a separate Python process, where each data item is a normal variable.
3. The log streams into the pane while Gurobi runs.
4. When the model ends, the helper compares the data items to their starting values. The add-in writes
   back only the ones that changed. If a result would replace a formula, you're asked first.

Your models, data items and solver settings are saved inside the workbook. Run logs are in
`~/Library/Application Support/GurobiStudio/runs/`.

## Turn the helper on and off

The helper doesn't start by itself unless you set up the login item below. If it's off, the pane shows
a red "Helper unreachable" badge. Closing Excel doesn't stop it.

**Check whether it's running.** Nothing printed means it's off.

```sh
pgrep -fl gurobistudio
```

**Turn it on.** This keeps running in that Terminal window, so leave the window open.

```sh
~/Desktop/SolverStudio/scripts/run-helper.sh
```

**Turn it off.** Press Ctrl+C in that window, or run this from any Terminal:

```sh
pkill -f gurobistudio
```

If a model is solving, stop it first with ⌘. in the pane. That lets Gurobi release its license token.
`pkill` also ends a running solve, and the result isn't written back.

### Start it at login (optional)

This keeps the helper running without a Terminal window. Set it up once:

```sh
cd ~/Desktop/SolverStudio
GS_LAUNCH_AGENT=1 scripts/install.sh
```

The login item restarts the helper if it's killed, so `pkill` won't keep it off. Use these instead:

```sh
# Off, until you turn it on again or next log in
launchctl bootout gui/$(id -u)/com.gurobistudio.helper

# On again
launchctl bootstrap gui/$(id -u) ~/Library/LaunchAgents/com.gurobistudio.helper.plist

# Remove the login item (run the "Off" command first)
rm ~/Library/LaunchAgents/com.gurobistudio.helper.plist
```

## More

- [docs/README.md](docs/README.md): install, examples, big models, troubleshooting
- [docs/TESTING.md](docs/TESTING.md): what has been checked in real Excel
- [docs/DECISIONS.md](docs/DECISIONS.md): why things are the way they are
- [CLAUDE.md](CLAUDE.md): developer commands and architecture
