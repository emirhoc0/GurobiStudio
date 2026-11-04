#!/bin/bash
# Removes the sideloaded manifest, the login item and the trusted dev certificate. Keeps your workbooks and repo.
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
rm -f "$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef/gurobistudio.manifest.xml"
launchctl bootout "gui/$(id -u)/com.gurobistudio.helper" 2>/dev/null || true
rm -f "$HOME/Library/LaunchAgents/com.gurobistudio.helper.plist"
pkill -f -- 'python.* -m gurobistudio$' || true
cd "$ROOT/add-in"
npx office-addin-dev-certs uninstall
echo "Removed. Helper data is still in ~/Library/Application Support/GurobiStudio."
