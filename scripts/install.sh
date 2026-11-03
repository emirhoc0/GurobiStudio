#!/bin/bash
# Safe to re-run. Options via env: GS_PORT (default 8765), GS_PLACEMENT (data|custom|home).
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
PORT="${GS_PORT:-8765}"
PLACEMENT="${GS_PLACEMENT:-data}"
WEF="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"

echo "==> Python environment"
[ -x "$ROOT/helper/.venv/bin/python" ] || python3.13 -m venv "$ROOT/helper/.venv"
"$ROOT/helper/.venv/bin/pip" install -q --disable-pip-version-check -r "$ROOT/helper/requirements.txt"
"$ROOT/helper/.venv/bin/python" -c "import gurobipy; print('gurobipy', gurobipy.gurobi.version())"

echo "==> Building the add-in"
cd "$ROOT/add-in"
npm ci --silent
npm run --silent build

echo "==> HTTPS certificates (macOS may ask for your password the first time)"
# 800 days: macOS rejects TLS certificates valid for more than 825.
npx office-addin-dev-certs install --days 800
chmod 600 "$HOME/.office-addin-dev-certs/localhost.key"

echo "==> Sideloading the manifest"
mkdir -p "$WEF"
python3 "$ROOT/scripts/make_manifest.py" "$PORT" "$PLACEMENT" > "$WEF/gurobistudio.manifest.xml"

echo "Done. Start the helper with scripts/run-helper.sh, then restart Excel."
