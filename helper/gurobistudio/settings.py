"""Helper-wide settings: which Python and license file runs use, and how many runs to keep."""

import json
import os
from pathlib import Path

DEFAULTS = {"python_path": "", "license_path": "", "keep_runs": 10}


def clean(values: dict) -> dict:
    """Raises ValueError with a message for the user."""
    out = {**DEFAULTS, **{k: values[k] for k in DEFAULTS if k in values}}
    python = str(out["python_path"]).strip()
    if python and not (Path(python).is_file() and os.access(python, os.X_OK)):
        raise ValueError(f"{python} is not an executable file.")
    license_path = str(out["license_path"]).strip()
    if license_path and not Path(license_path).is_file():
        raise ValueError(f"{license_path} is not a file.")
    keep = out["keep_runs"]
    if not isinstance(keep, int) or isinstance(keep, bool) or not 1 <= keep <= 1000:
        raise ValueError("Keep runs must be a whole number from 1 to 1000.")
    return {"python_path": python, "license_path": license_path, "keep_runs": keep}


def load(path: Path) -> dict:
    try:
        return clean(json.loads(path.read_text()))
    except (OSError, ValueError):
        return dict(DEFAULTS)  # a moved Python shouldn't stop the helper; Settings shows defaults


def child_env(settings: dict) -> dict:
    env = dict(os.environ)
    if settings["license_path"]:
        env["GRB_LICENSE_FILE"] = settings["license_path"]
    return env
