#!/usr/bin/env python3
"""Fills add-in/manifest.template.xml. Usage: make_manifest.py PORT [data|custom|home] > out.xml

Run with the helper venv's Python. The output holds the page key, so write it to a private file.
"""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "helper"))

from gurobistudio.main import default_data_dir  # noqa: E402
from gurobistudio.security import load_or_create_secret  # noqa: E402

port = sys.argv[1]
placement = sys.argv[2] if len(sys.argv) > 2 else "data"
OPEN = {
    "data": '<OfficeTab id="TabData">',
    "custom": '<CustomTab id="GS.Tab">',
    "home": '<OfficeTab id="TabHome">',
}
CLOSE = {
    "data": "</OfficeTab>",
    "custom": '<Label resid="Group.Label"/></CustomTab>',
    "home": "</OfficeTab>",
}
template = (ROOT / "add-in" / "manifest.template.xml").read_text()
key = load_or_create_secret(default_data_dir(), "page_key")
print(
    template.replace("{{PAGE}}", f"https://localhost:{port}/taskpane.html?k={key}")
    .replace("{{BASE}}", f"https://localhost:{port}")
    .replace("{{TAB_OPEN}}", OPEN[placement])
    .replace("{{TAB_CLOSE}}", CLOSE[placement])
)
