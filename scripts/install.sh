#!/bin/bash
# Safe to re-run. Options via env: GS_PORT (default 8765), GS_PLACEMENT (data|custom|home),
# GS_LAUNCH_AGENT=1 (start the helper at login and keep it running).
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
PORT="${GS_PORT:-8765}"
PLACEMENT="${GS_PLACEMENT:-data}"
WEF="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"

DATA="${GS_DATA_DIR:-$HOME/Library/Application Support/GurobiStudio}"

echo "==> Python environment"
[ -x "$ROOT/helper/.venv/bin/python" ] || python3.13 -m venv "$ROOT/helper/.venv"
"$ROOT/helper/.venv/bin/pip" install -q --disable-pip-version-check -r "$ROOT/helper/requirements.txt"
"$ROOT/helper/.venv/bin/python" -c "import gurobipy; print('gurobipy', gurobipy.gurobi.version())"

echo "==> Building the add-in"
cd "$ROOT/add-in"
# npm ci occasionally loses a race deleting node_modules (ENOTEMPTY); a retry has cleared it.
npm ci --silent || npm ci --silent
npm run --silent build

echo "==> HTTPS certificates (macOS may ask for your password the first time)"
# 800 days: macOS rejects TLS certificates valid for more than 825.
npx office-addin-dev-certs install --days 800
chmod 600 "$HOME/.office-addin-dev-certs/localhost.key"

echo "==> Sideloading the manifest"
mkdir -p "$WEF"
# Holds the page key, so only this account may read it.
(umask 077; "$ROOT/helper/.venv/bin/python" "$ROOT/scripts/make_manifest.py" "$PORT" "$PLACEMENT" \
  > "$WEF/gurobistudio.manifest.xml")
chmod 600 "$WEF/gurobistudio.manifest.xml"
mkdir -p "$DATA"
echo "$PORT" > "$DATA/port" # the helper reads this, so it always matches the manifest

AGENT=com.gurobistudio.helper
PLIST="$HOME/Library/LaunchAgents/$AGENT.plist"
if [ "${GS_LAUNCH_AGENT:-0}" = 1 ]; then
  echo "==> Starting the helper at login"
  mkdir -p "$(dirname "$PLIST")"
  launchctl bootout "gui/$(id -u)/$AGENT" 2>/dev/null || true
  "$ROOT/helper/.venv/bin/python" "$ROOT/scripts/make_plist.py" "$PLIST" "$AGENT" \
    "$ROOT/scripts/run-helper.sh" "$DATA/helper.log"
  launchctl bootstrap "gui/$(id -u)" "$PLIST"
  echo "Done. The helper is running and starts at login. Restart Excel."
elif [ -f "$PLIST" ]; then
  launchctl kickstart -k "gui/$(id -u)/$AGENT" # picks up the code just installed
  echo "Done. The helper was restarted. Restart Excel."
else
  echo "Done. Start the helper with scripts/run-helper.sh, then restart Excel."
fi
