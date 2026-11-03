#!/usr/bin/env python3
"""Fills add-in/manifest.template.xml. Usage: make_manifest.py PORT [data|custom|home] > out.xml"""

import sys
from pathlib import Path

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
template = (Path(__file__).parent.parent / "add-in" / "manifest.template.xml").read_text()
print(
    template.replace("{{BASE}}", f"https://localhost:{port}")
    .replace("{{TAB_OPEN}}", OPEN[placement])
    .replace("{{TAB_CLOSE}}", CLOSE[placement])
)
