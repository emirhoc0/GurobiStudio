"""Runs one model in its own process: python -u -m gurobistudio.runner RUN_DIR

Results go to result.json, never stdout, so a model's print can't corrupt them.
"""

import copy
import json
import os
import sys
import traceback
from pathlib import Path

from .wire import NotSerializable, decode, encode

MODEL_FILE = "model.py"
STATUS_NAMES = {
    2: "Optimal",
    3: "Infeasible",
    4: "Infeasible or unbounded",
    5: "Unbounded",
    9: "Time limit",
    11: "Interrupted",
}
HINTS = {
    "clr": "IronPython features such as clr are not supported.",
    "System": "IronPython features such as System are not supported.",
}


def error_info(exc: BaseException) -> dict:
    if isinstance(exc, SyntaxError):
        hint = None
        if "Missing parentheses in call to 'print'" in str(exc):
            hint = "This is Python 2 syntax. Add parentheses: print('text')."
        return {
            "type": type(exc).__name__,
            "message": exc.msg,
            "line": exc.lineno,
            "traceback": "".join(traceback.format_exception_only(exc)),
            "hint": hint,
        }
    frames = traceback.extract_tb(exc.__traceback__)
    model_frames = [f for f in frames if f.filename == MODEL_FILE]
    return {
        "type": type(exc).__name__,
        "message": str(exc),
        "line": model_frames[-1].lineno if model_frames else None,
        "traceback": "".join(traceback.format_exception(exc)),
        "hint": HINTS.get(getattr(exc, "name", None)),
    }


def is_same(final, snapshot) -> bool:
    try:
        return bool(final == snapshot)
    except (ValueError, TypeError):  # e.g. numpy arrays have no single truth value
        return False


def model_status(namespace: dict) -> tuple[str | None, float | None]:
    gurobipy = sys.modules.get("gurobipy")
    if gurobipy is None:
        return None, None
    models = [v for v in namespace.values() if isinstance(v, gurobipy.Model)]
    if not models:
        return None, None
    model = models[-1]
    status = STATUS_NAMES.get(model.Status, f"Status {model.Status}")
    try:
        return status, model.ObjVal
    except (AttributeError, gurobipy.GurobiError):
        return status, None


def collect_changes(spec_items, namespace, snapshot) -> tuple[list[dict], list[str]]:
    changed, warnings = [], []
    for item in spec_items:
        name = item["name"]
        if name not in namespace:
            warnings.append(f"{name} was deleted by the model, so its cells were left alone.")
        elif is_same(namespace[name], snapshot[name]):
            pass
        elif item.get("readOnly"):
            warnings.append(f"{name} changed but is read-only, so it was not written back.")
        else:
            try:
                changed.append({"name": name, "value": encode(namespace[name])})
            except NotSerializable as e:
                warnings.append(f"{name} holds a value Excel can't store ({e}); not written back.")
    return changed, warnings


def run(run_dir: Path) -> dict:
    spec = json.loads((run_dir / "input.json").read_text())
    source = spec["model"]
    (run_dir / MODEL_FILE).write_text(source)
    os.chdir(run_dir)

    items = {i["name"]: decode(i["value"]) for i in spec["items"]}
    snapshot = copy.deepcopy(items)

    sys.path.insert(0, str(Path(__file__).parent / "shim"))
    import SolverStudio

    SolverStudio._install(items)

    # Same objects as in the SolverStudio module, so in-place edits and reassignment both show up.
    namespace = {"__name__": "__main__", "__file__": MODEL_FILE, **items}
    namespace["Application"] = SolverStudio.Application

    try:
        if spec.get("params"):
            import gurobipy

            for key, value in spec["params"].items():
                gurobipy.setParam(key, value)
        exec(compile(source, MODEL_FILE, "exec"), namespace)  # noqa: S102  # running it is the point
    except SystemExit as e:
        if e.code not in (None, 0):
            return {"ok": False, "error": error_info(e)}
    except Exception as e:  # noqa: BLE001  # any model failure becomes an error card
        return {"ok": False, "error": error_info(e)}

    status, objective = model_status(namespace)
    changed, warnings = collect_changes(spec["items"], namespace, snapshot)
    changed_names = {c["name"] for c in changed}
    names = [i["name"] + ("*" if i["name"] in changed_names else "") for i in spec["items"]]
    print("Data items: " + (", ".join(names) or "none"))
    return {
        "ok": True,
        "status": status,
        "objective": objective,
        "changed": changed,
        "warnings": warnings,
    }


if __name__ == "__main__":
    run_dir = Path(sys.argv[1]).resolve()
    result = run(run_dir)
    sys.stdout.flush()
    (run_dir / "result.json").write_text(json.dumps(result))
