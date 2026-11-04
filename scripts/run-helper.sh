#!/bin/bash
set -euo pipefail
cd "$(dirname "$0")/../helper"
if [ ! -x .venv/bin/python ] || [ ! -f "$HOME/.office-addin-dev-certs/localhost.key" ]; then
  echo "GurobiStudio isn't installed yet. Run scripts/install.sh first." >&2
  exit 1
fi
exec .venv/bin/python -m gurobistudio
