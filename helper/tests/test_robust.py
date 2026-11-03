"""A solved model's results must survive whatever the model leaves behind, and the helper must not jam."""

import signal
import sys
import time
import types
from pathlib import Path

import pytest
from test_runs import HEADER, api, changed, events, item, solve, start  # noqa: F401

from gurobistudio import runs as runs_module
from gurobistudio.runner import model_status
from gurobistudio.runs import RunManager
from gurobistudio.wire import encode

RUN_BODY = {"items": [], "params": {}, "workbook_id": "wb", "approved": True}


def wait_finished(run):
    deadline = time.time() + 20
    while not run.finished and time.time() < deadline:
        time.sleep(0.01)
    assert run.finished


def test_gurobi_vars_over_a_filled_item_keep_the_other_results(api):  # noqa: F811
    model = HEADER + "m = Model()\nx = m.addVars(1)\nm.update()\nkeep = 7\n"
    items = [item("x", {"__dict__": [[0, 5.0]]}), item("keep", 1)]
    final, _ = solve(api, model, items)
    assert final["type"] == "result"
    assert changed(final) == {"keep": 7}
    assert any(w.startswith("x holds a value") for w in final["warnings"])


def test_disposed_model_keeps_the_results(api):  # noqa: F811
    final, _ = solve(api, HEADER + "m = Model()\nm.dispose()\nout = 2\n", [item("out", 1)])
    assert final["type"] == "result" and changed(final) == {"out": 2}


def test_a_value_that_fails_to_encode_only_loses_that_item(api):  # noqa: F811
    model = (
        "class Bad:\n"
        "    def item(self):\n"
        "        raise ValueError('only size-1 arrays')\n"
        "out = Bad()\n"
        "keep = 2\n"
    )
    final, _ = solve(api, model, [item("out", 1), item("keep", 1)])
    assert final["type"] == "result" and changed(final) == {"keep": 2}
    assert any(w.startswith("out holds a value") for w in final["warnings"])


def test_arrays_encode_through_tolist():
    class Array:
        def item(self):
            raise ValueError("only size-1 arrays")

        def tolist(self):
            return [[1.0, 2.0], [3.0, float("inf")]]

    assert encode(Array()) == [[1.0, 2.0], [3.0, "inf"]]
    assert encode([Array()]) == [[[1.0, 2.0], [3.0, "inf"]]]


@pytest.mark.parametrize("bad", [float("inf"), float("-inf"), float("nan")])
def test_non_finite_objective_is_not_sent(monkeypatch, bad):
    class Model:
        Status = 2
        ObjVal = bad

    monkeypatch.setitem(
        sys.modules, "gurobipy", types.SimpleNamespace(Model=Model, GurobiError=RuntimeError)
    )
    assert model_status({"m": Model()}) == ("Optimal", None)


def test_unoptimized_model_has_a_status_name(api):  # noqa: F811
    final, _ = solve(api, HEADER + "m = Model()\n")
    assert final["status"] == "Loaded"


def test_a_failing_pump_still_finishes_the_run(tmp_path, monkeypatch):
    def boom(*_):
        raise OSError("disk full")

    monkeypatch.setattr(RunManager, "_final_events", boom)
    manager = RunManager(tmp_path)
    run = manager.start({"model": "x = 1\n", "items": [], "params": {}})
    wait_finished(run)
    last = list(run.events)[-1]
    assert last["type"] == "error" and "disk full" in last["error"]["message"]
    assert not manager.busy()


def test_busy_refusal_names_the_running_run(api):  # noqa: F811
    run_id = start(api, "import time\ntime.sleep(30)\n")
    r = api.post("/api/run", json={**RUN_BODY, "model": "x=1"})
    assert r.status_code == 409 and r.json() == {"busy": True, "run_id": run_id}
    api.post(f"/api/run/{run_id}/cancel")
    events(api, run_id)


def test_missing_python_is_reported(api, tmp_path):  # noqa: F811
    python = tmp_path / "python3"
    python.write_text("#!/bin/sh\n")
    python.chmod(0o755)
    assert api.put("/api/settings", json={"python_path": str(python)}).status_code == 200
    python.unlink()
    r = api.post("/api/run", json={**RUN_BODY, "model": "x=1"})
    assert r.status_code == 500 and str(python) in r.json()["error"]
    assert api.post("/api/run", json={**RUN_BODY, "model": "x=1"}).status_code == 500  # not busy


def test_a_stray_folder_in_runs_is_ignored(tmp_path):
    (tmp_path / "untitled folder").mkdir()
    manager = RunManager(tmp_path)
    run = manager.start({"model": "x = 1\n", "items": [], "params": {}})
    wait_finished(run)
    assert manager.latest_dir() == run.dir
    assert (tmp_path / "untitled folder").exists()


def test_no_sigkill_after_the_run_has_ended(tmp_path, monkeypatch):
    sent = []
    monkeypatch.setattr(runs_module, "KILL_AFTER_S", 0.3)
    original = RunManager._signal

    def spy(pgid, sig):
        sent.append(sig)
        original(pgid, sig)

    monkeypatch.setattr(RunManager, "_signal", staticmethod(spy))
    manager = RunManager(tmp_path)
    run = manager.start({"model": "import time\ntime.sleep(30)\n", "items": [], "params": {}})
    time.sleep(0.5)  # let the interpreter reach the sleep, so SIGINT ends it at once
    manager.cancel(run)
    wait_finished(run)
    time.sleep(0.5)
    assert sent == [signal.SIGINT]


def test_non_ascii_survives_a_c_locale(api, monkeypatch):  # noqa: F811
    monkeypatch.setenv("LC_ALL", "C")
    monkeypatch.setenv("PYTHONUTF8", "0")
    monkeypatch.setenv("PYTHONCOERCECLOCALE", "0")
    final, log = solve(api, "print('Üretim şğı')\nout = 'Kâr'\n", [item("out", "")])
    assert final["type"] == "result" and "Üretim şğı" in log
    assert changed(final) == {"out": "Kâr"}


@pytest.mark.parametrize("name", ["WLSSECRET", "wls_secret", "WLS-Secret", "wlsaccessid"])
def test_credentials_stay_out_of_logs_and_error_cards(api, name):  # noqa: F811
    secret = "s3cr3tvalue"
    model = f"opts = {{'{name}': '{secret}'}}\nprint(opts)\nraise ValueError(str(opts))\n"
    run_id = start(api, model)
    evs = events(api, run_id)
    assert evs[-1]["type"] == "error" and evs[-1]["error"]["type"] == "ValueError"
    assert secret not in str(evs)
    run_dir = api.get(f"/api/run/{run_id}/files").json()["dir"]
    assert secret not in Path(run_dir, "log.txt").read_text()
