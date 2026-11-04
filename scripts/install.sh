#!/bin/bash
# Safe to re-run. Options via env: GS_PORT (default 8765), GS_PLACEMENT (data|custom|home),
# GS_LAUNCH_AGENT=0 (don't start the helper at login) or =1 (start it even from a protected folder).
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
PORT="${GS_PORT:-8765}"
PLACEMENT="${GS_PLACEMENT:-data}"
CONTAINER="$HOME/Library/Containers/com.microsoft.Excel"
WEF="$CONTAINER/Data/Documents/wef"

DATA="${GS_DATA_DIR:-$HOME/Library/Application Support/GurobiStudio}"
AGENT=com.gurobistudio.helper
PLIST="$HOME/Library/LaunchAgents/$AGENT.plist"

fail() {
  echo "ERROR: $1" >&2
  exit 1
}

echo "==> Checking requirements"
[ "$(uname)" = Darwin ] || fail "GurobiStudio only runs on macOS."
[ -d "/Applications/Microsoft Excel.app" ] ||
  fail "Microsoft Excel isn't in /Applications. Install it, open it once, then run this again."
[ -d "$CONTAINER" ] || fail "Excel has never been opened on this account. Open it once, quit it, then run this again."
[ -x "$ROOT/helper/.venv/bin/python" ] || command -v python3.13 >/dev/null ||
  fail "Python 3.13 isn't installed. Run: brew install python@3.13"
command -v node >/dev/null || fail "Node isn't installed. Run: brew install node"
# Vite's minimum.
node -e 'const [a, b] = process.versions.node.split(".").map(Number); process.exit(a > 22 || (a === 22 && b >= 12) ? 0 : 1)' ||
  fail "Node $(node -v) is too old, 22.12 or newer is needed. Run: brew upgrade node"
# macOS may ask to let Terminal "access data from other apps" here. That's Excel's add-in folder.
{ mkdir -p "$WEF" && touch "$WEF/.gs-write-test" && rm "$WEF/.gs-write-test"; } 2>/dev/null ||
  fail "macOS blocked access to Excel's add-in folder. Allow your Terminal app in System Settings > Privacy & Security (App Management, or Files and Folders), then run this again."

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
# The tool exits 0 even when it couldn't install the certificate, and Excel then can't load the pane.
security verify-cert -c "$HOME/.office-addin-dev-certs/localhost.crt" -p ssl -s localhost -q >/dev/null 2>&1 ||
  fail "macOS doesn't trust the local certificate, so Excel couldn't load the add-in. Run this again and enter your password when asked."

echo "==> Sideloading the manifest"
# Holds the page key, so only this account may read it.
(umask 077; "$ROOT/helper/.venv/bin/python" "$ROOT/scripts/make_manifest.py" "$PORT" "$PLACEMENT" \
  > "$WEF/gurobistudio.manifest.xml")
chmod 600 "$WEF/gurobistudio.manifest.xml"
mkdir -p "$DATA"
echo "$PORT" > "$DATA/port" # the helper reads this, so it always matches the manifest

# The login item is the default, because the pane and the ribbon buttons can't load without the
# helper. launchd jobs may be blocked from these folders, so there it takes an explicit opt-in.
LOGIN_ITEM="${GS_LAUNCH_AGENT:-}"
PROTECTED=0
if [ -z "$LOGIN_ITEM" ]; then
  LOGIN_ITEM=1
  case "$ROOT/" in
    "$HOME"/Desktop/* | "$HOME"/Documents/* | "$HOME"/Downloads/*)
      PROTECTED=1
      LOGIN_ITEM=0
      ;;
  esac
  [ ! -f "$PLIST" ] || LOGIN_ITEM=1 # already chosen on an earlier install
fi

echo "==> Restarting the helper"
launchctl bootout "gui/$(id -u)/$AGENT" 2>/dev/null || true
# Picks up the code just installed. The pattern leaves a running model's process alone.
if pkill -f -- 'python.* -m gurobistudio$'; then sleep 1; fi
if lsof -ti "tcp:$PORT" -sTCP:LISTEN >/dev/null; then
  fail "Another program is using port $PORT. Run this again with another port: GS_PORT=9000 scripts/install.sh"
fi

if [ "$LOGIN_ITEM" = 1 ]; then
  mkdir -p "$(dirname "$PLIST")"
  "$ROOT/helper/.venv/bin/python" "$ROOT/scripts/make_plist.py" "$PLIST" "$AGENT" \
    "$ROOT/scripts/run-helper.sh" "$DATA/helper.log"
  # launchd can refuse this for a moment after a bootout.
  for _ in 1 2 3 4 5; do
    if launchctl bootstrap "gui/$(id -u)" "$PLIST" 2>/dev/null; then break; fi
    sleep 1
  done
  for _ in $(seq 1 20); do
    if curl -sk -o /dev/null --max-time 2 "https://127.0.0.1:$PORT/"; then break; fi
    sleep 1
  done
  curl -sk -o /dev/null --max-time 2 "https://127.0.0.1:$PORT/" ||
    fail "The helper didn't start. See \"$DATA/helper.log\"."
  DONE="Done. The helper is running and starts at login. Restart Excel."
else
  rm -f "$PLIST"
  [ "$PROTECTED" = 0 ] ||
    echo "Note: no login item, because macOS may block one in Desktop, Documents and Downloads. Clone to ~/GurobiStudio to get one."
  DONE="Done. Start the helper with scripts/run-helper.sh, then restart Excel."
fi

if [ ! -f "${GRB_LICENSE_FILE:-$HOME/gurobi.lic}" ] && [ ! -f /Library/gurobi/gurobi.lic ]; then
  echo "Note: no Gurobi license found, so models are limited to 2,000 variables and 2,000 constraints."
  echo "      See \"Get your own Gurobi license\" in the README."
fi
echo "$DONE"
