#!/bin/bash
# Removes the sideloaded manifest and the trusted dev certificate. Keeps your workbooks and repo.
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
rm -f "$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef/gurobistudio.manifest.xml"
cd "$ROOT/add-in"
npx office-addin-dev-certs uninstall
echo "Removed. Helper data is still in ~/Library/Application Support/GurobiStudio."
