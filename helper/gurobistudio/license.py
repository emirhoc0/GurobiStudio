"""Check License: reports only derived facts, never raw log lines, so nothing private can leak."""

import json
import re
import subprocess
import tempfile

from .runs import CREDENTIAL_LINE, REGISTERED_TO

TIMEOUT_S = 60
RESULT_PREFIX = "GS_LICENSE_RESULT "
# 2,001 variables: the definitive test, since the size-limited license refuses more than 2,000.
SCRIPT = f"""
import json
out = {{}}
try:
    import gurobipy as gp
    tiny = gp.Model()
    tiny.addVar()
    tiny.optimize()
    out["expiry"] = tiny.LicenseExpiration
    big = gp.Model()
    big.Params.OutputFlag = 0
    big.addVars(2001)
    try:
        big.optimize()
        out["size_ok"] = True
    except gp.GurobiError as e:
        out["size_ok"] = False
        out["size_error"] = str(e)
except Exception as e:
    out["error"] = str(e)
print({RESULT_PREFIX!r} + json.dumps(out))
"""

TYPES = (
    ("restricted", "Restricted"),
    ("academic", "Academic"),
    ("evaluation", "Evaluation"),
    ("trial", "Trial"),
    ("commercial", "Commercial"),
    ("named[- ]user", "Named-user"),
    ("floating", "Floating"),
    ("wls", "WLS"),
)
OFFLINE = re.compile(
    r"could not (connect|resolve|reach)|unable to (connect|resolve|reach)|failed to connect"
    r"|cannot connect|no route|network|timed? ?out|name resolution|connection (refused|reset)"
    r"|ssl|unreachable",
    re.IGNORECASE,
)
LOG_EXPIRY = re.compile(r"expires?\s+(?:on\s+)?(\d{4}-\d{2}-\d{2})", re.IGNORECASE)
NO_EXPIRY = {0, 99991231}


def _scrub(text: str) -> str:
    if CREDENTIAL_LINE.search(text):
        return "Gurobi reported a license error."
    return REGISTERED_TO.sub(r"\1(hidden)", text)


def _format_expiry(value, log: str) -> str | None:
    if isinstance(value, int) and value not in NO_EXPIRY:
        text = str(value)
        return f"{text[:4]}-{text[4:6]}-{text[6:]}"
    match = LOG_EXPIRY.search(log)
    return match.group(1) if match else None


def parse_license(log: str, facts: dict) -> dict:
    """Matches loosely, because Gurobi's wording changes between versions."""
    license_lines = "\n".join(line for line in log.splitlines() if "licen" in line.lower())
    license_type = "Unknown"
    for pattern, name in TYPES:
        if re.search(rf"\b{pattern}\b", license_lines, re.IGNORECASE):
            license_type = name
            break

    error = facts.get("error")
    if error:
        offline = bool(OFFLINE.search(error) or OFFLINE.search(license_lines))
        message = (
            "Could not reach Gurobi's license service. A WLS license needs internet access."
            if offline
            else _scrub(error)
        )
        return {
            "ok": False,
            "type": license_type,
            "expiry": None,
            "restricted": False,
            "offline": offline,
            "message": message,
        }

    restricted = facts.get("size_ok") is False or license_type == "Restricted"
    if restricted:
        license_type = "Restricted"
        message = (
            "Only the size-limited license is active, so models over 2,000 variables will fail. "
            "Gurobi did not find your license file."
        )
    else:
        message = "The license works, including for models over the size limit."
    return {
        "ok": True,
        "type": license_type,
        "expiry": _format_expiry(facts.get("expiry"), license_lines),
        "restricted": restricted,
        "offline": False,
        "message": message,
    }


def check_license(python: str, env: dict) -> dict:
    with tempfile.TemporaryDirectory() as cwd:  # Gurobi may drop a gurobi.log in its working folder
        try:
            done = subprocess.run(
                [python, "-u", "-c", SCRIPT],
                cwd=cwd,
                env=env,
                capture_output=True,
                text=True,
                errors="replace",
                timeout=TIMEOUT_S,
                check=False,
            )
        except subprocess.TimeoutExpired:
            return parse_license("", {"error": "The license check timed out."})
    lines = done.stdout.splitlines()
    results = [line for line in lines if line.startswith(RESULT_PREFIX)]
    if not results:
        return parse_license(done.stdout, {"error": "The license check did not finish."})
    log = "\n".join(line for line in lines if not line.startswith(RESULT_PREFIX))
    return parse_license(log, json.loads(results[-1][len(RESULT_PREFIX) :]))
