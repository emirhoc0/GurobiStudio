"""Writes the login item. Usage: make_plist.py OUT LABEL SCRIPT LOG"""

import plistlib
import sys

out, label, script, log = sys.argv[1:]
with open(out, "wb") as f:
    plistlib.dump(
        {
            "Label": label,
            "ProgramArguments": ["/bin/bash", script],
            "RunAtLoad": True,
            "KeepAlive": {"SuccessfulExit": False},
            "StandardOutPath": log,
            "StandardErrorPath": log,
        },
        f,
    )
