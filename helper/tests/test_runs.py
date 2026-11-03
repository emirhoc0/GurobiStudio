"""Runs real models through the real API, with real Gurobi where the model imports it."""

import json
import os
import time
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from gurobistudio.main import create_app
from gurobistudio.runner import is_same
from gurobistudio.runs import status_from_log

HEADER = "from SolverStudio import *\nfrom gurobipy import *\n"
BASE = "https://localhost:8765"


@pytest.fixture
def api(tmp_path):
    dist = tmp_path / "dist"
    dist.mkdir()
    (dist / "taskpane.html").write_text("x")
    client = TestClient(create_app(data_dir=tmp_path / "data", dist_dir=dist), base_url=BASE)
    client.headers["Authorization"] = f"Bearer {(tmp_path / 'data' / 'token').read_text()}"
    return client


def start(api, model, items=(), params=None):
    r = api.post(
        "/api/run",
        json={
            "model": model,
            "items": list(items),
            "params": params or {},
            "workbook_id": "wb",
            "approved": True,
        },
    )
    assert r.status_code == 200, r.text
    return r.json()["run_id"]


def events(api, run_id):
    text = api.get(f"/api/run/{run_id}/events").text
    return [json.loads(c[6:]) for c in text.split("\n\n") if c.startswith("data: ")]


def solve(api, model, items=(), params=None):
    evs = events(api, start(api, model, items, params))
    log = [e["line"] for e in evs if e["type"] == "log"]
    return evs[-1], log


def changed(final):
    return {c["name"]: c["value"] for c in final["changed"]}


def item(name, value, **extra):
    return {"name": name, "type": "x", "value": value, **extra}


def test_in_place_mutation_is_detected(api):
    model = "from SolverStudio import *\nout['a'] = 5\nout['b'] = 6.5\n"
    final, log = solve(api, model, [item("out", {"__dict__": [["a", 1]]}), item("keep", 1)])
    assert final["ok"] and changed(final) == {"out": {"__dict__": [["a", 5], ["b", 6.5]]}}
    assert "Data items: out*, keep" in log


def test_reassignment_is_detected(api):
    final, log = solve(api, "from SolverStudio import *\nout = {'a': 9}\n", [item("out", None)])
    assert changed(final) == {"out": {"__dict__": [["a", 9]]}}
    assert "Data items: out*" in log


def test_items_exist_without_the_import(api):
    final, _ = solve(api, "b = a + 1\n", [item("a", 1), item("b", 0)])
    assert changed(final) == {"b": 2}


def test_tuple_keys_reach_the_model(api):
    model = "total = sum(costs[w, b] for w, b in costs)\nfirst = arcs[0][1]\n"
    costs = {"__dict__": [[["w", "b"], 2], [["x", "y"], 3]]}
    final, _ = solve(
        api,
        model,
        [item("costs", costs), item("arcs", [["p", "q"]]), item("total", 0), item("first", 0)],
    )
    assert changed(final) == {"total": 5, "first": "q"}


def test_read_only_item_is_not_written(api):
    final, _ = solve(api, "a = 2\n", [item("a", 1, readOnly=True)])
    assert changed(final) == {} and "read-only" in final["warnings"][0]


def test_unserializable_value_warns(api):
    final, _ = solve(api, "a = object()\n", [item("a", 1)])
    assert changed(final) == {} and "can't store" in final["warnings"][0]


def test_deleted_item_warns(api):
    final, _ = solve(api, "del a\n", [item("a", 1)])
    assert changed(final) == {} and "deleted" in final["warnings"][0]


def test_real_gurobi_solves_and_writes(api):
    model = HEADER + (
        "m = Model('t')\n"
        "x = m.addVar(ub=capacity, name='x')\n"
        "m.setObjective(2 * x, GRB.MAXIMIZE)\n"
        "m.optimize()\n"
        "result = x.x\n"
        "if m.status == GRB.status.OPTIMAL:\n"
        "    status = 'Optimal'\n"
    )
    final, log = solve(
        api, model, [item("capacity", 10), item("result", None), item("status", None)]
    )
    assert final["type"] == "result" and final["status"] == "Optimal"
    assert final["objective"] == 20
    assert changed(final) == {"result": 10, "status": "Optimal"}
    assert any("Optimal objective" in line for line in log)
    assert "Data items: capacity, result*, status*" in log


def test_no_credential_lines_in_stream_or_files(api):
    run_id = start(api, HEADER + "m = Model()\nm.optimize()\n")
    evs = events(api, run_id)
    text = "\n".join(e.get("line", "") for e in evs)
    run_dir = api.get(f"/api/run/{run_id}/files").json()["dir"]
    on_disk = Path(run_dir, "log.txt").read_text()
    for blob in (text, on_disk):
        assert "wlssecret" not in blob.lower() and "wlsaccessid" not in blob.lower()
        assert "licenseid" not in blob.lower()


def test_infeasible_reports_status_and_writes_only_what_was_assigned(api):
    model = HEADER + (
        "m = Model()\n"
        "x = m.addVar(ub=1)\n"
        "m.addConstr(x >= 2)\n"
        "m.optimize()\n"
        "status = 'no solution'\n"
    )
    final, _ = solve(api, model, [item("status", None), item("result", 1)])
    assert final["status"] == "Infeasible" and final["objective"] is None
    assert changed(final) == {"status": "no solution"}


def test_exception_reports_line_and_writes_nothing(api):
    model = "from SolverStudio import *\nout = 5\n\ndef f():\n    raise ValueError('boom')\nf()\n"
    final, _ = solve(api, model, [item("out", 1)])
    assert final["type"] == "error"
    assert final["error"]["line"] == 5 and final["error"]["type"] == "ValueError"
    assert "boom" in final["error"]["traceback"]


def test_python2_print_gets_a_hint(api):
    final, _ = solve(api, "x = 1\nprint 'hello'\n")
    assert final["type"] == "error" and final["error"]["line"] == 2
    assert "Add parentheses" in final["error"]["hint"]


def test_application_object_fails_clearly(api):
    final, _ = solve(api, "from SolverStudio import *\nApplication.ActiveSheet\n")
    assert (
        final["type"] == "error"
        and "Application object is not supported" in final["error"]["message"]
    )


def test_ironpython_import_fails_clearly(api):
    final, _ = solve(api, "import clr\n")
    assert "IronPython" in final["error"]["hint"]


def test_gurobi_params_are_applied(api):
    final, log = solve(api, HEADER + "m = Model()\nm.optimize()\n", params={"TimeLimit": 7})
    assert final["ok"] and any("TimeLimit to value 7" in line for line in log)


def test_stop_kills_whole_group_quickly(api):
    # The grandchild ignores SIGINT, so this also proves the SIGKILL escalation.
    model = (
        "import subprocess, sys, pathlib\n"
        "p = subprocess.Popen([sys.executable, '-c', "
        "'import signal,time; signal.signal(signal.SIGINT, signal.SIG_IGN); time.sleep(60)'], "
        "start_new_session=False)\n"
        "pathlib.Path('child.pid').write_text(str(p.pid))\n"
        "print('started', flush=True)\n"
        "import time\n"
        "time.sleep(60)\n"
    )
    run_id = start(api, model)
    run_dir = api.get(f"/api/run/{run_id}/files").json()["dir"]
    pid_file = os.path.join(run_dir, "child.pid")
    deadline = time.time() + 10
    while not os.path.exists(pid_file) and time.time() < deadline:
        time.sleep(0.05)
    t0 = time.time()
    assert api.post(f"/api/run/{run_id}/cancel").status_code == 200
    evs = events(api, run_id)
    assert evs[-1] == {"type": "status", "state": "stopped"}
    time.sleep(3.5)
    assert time.time() - t0 < 8
    with pytest.raises(ProcessLookupError):
        os.kill(int(Path(pid_file).read_text()), 0)


def test_stop_during_a_real_mip(api):
    model = HEADER + (
        "import random\n"
        "random.seed(1)\n"
        "m = Model()\n"
        "x = m.addVars(300, vtype=GRB.BINARY)\n"
        "for _ in range(100):\n"
        "    m.addConstr(sum(random.randint(1,99)*x[i] for i in random.sample(range(300),60)) <= 900)\n"
        "m.setObjective(sum(random.randint(1,99)*x[i] for i in range(300)), GRB.MAXIMIZE)\n"
        "m.optimize()\n"
        "out = 1\n"
    )
    run_id = start(api, model, [item("out", 0)])
    log = os.path.join(api.get(f"/api/run/{run_id}/files").json()["dir"], "log.txt")
    deadline = time.time() + 30
    while time.time() < deadline:
        if os.path.exists(log) and "Nodes" in Path(log).read_text():
            break
        time.sleep(0.1)
    t0 = time.time()
    api.post(f"/api/run/{run_id}/cancel")
    evs = events(api, run_id)
    assert evs[-1] == {"type": "status", "state": "stopped"}
    assert time.time() - t0 < 5


def test_trust_required_then_remembered_then_reprompted(api):
    body = {"model": "x = 1\n", "items": [], "params": {}, "workbook_id": "wb"}
    assert api.post("/api/run", json=body).json() == {"trust_required": True}
    run_id = api.post("/api/run", json={**body, "approved": True}).json()["run_id"]
    events(api, run_id)
    r = api.post("/api/run", json=body)
    assert r.status_code == 200
    events(api, r.json()["run_id"])
    assert api.post("/api/run", json={**body, "model": "x = 2\n"}).status_code == 409
    assert api.post("/api/run", json={**body, "workbook_id": "other"}).status_code == 409


def test_second_run_while_busy_is_refused(api):
    run_id = start(api, "import time\ntime.sleep(30)\n")
    r = api.post(
        "/api/run",
        json={"model": "x=1", "items": [], "params": {}, "workbook_id": "wb", "approved": True},
    )
    assert r.status_code == 409 and r.json()["busy"] is True
    api.post(f"/api/run/{run_id}/cancel")
    events(api, run_id)


def test_old_runs_are_pruned(api):
    last = None
    for _ in range(12):
        last = start(api, "x = 1\n")
        events(api, last)
    run_dir = api.get(f"/api/run/{last}/files").json()["dir"]
    assert len(os.listdir(os.path.dirname(run_dir))) == 10


def test_unknown_run_is_404(api):
    assert api.get("/api/run/../../etc/files").status_code == 404
    assert api.post("/api/run/nope/cancel").status_code == 404


def test_registered_email_is_hidden_in_stream_and_file(api):
    run_id = start(api, 'print("Academic license - registered to someone@example.com")\n')
    evs = events(api, run_id)
    lines = [e["line"] for e in evs if e["type"] == "log"]
    assert "Academic license - registered to (hidden)" in lines
    files = api.get(f"/api/run/{run_id}/files").json()
    assert "someone@example.com" not in (Path(files["dir"]) / "log.txt").read_text()


def test_bad_gurobi_param_gets_an_error_card(api):
    final, _ = solve(api, "x = 1\n", params={"NoSuchParameter": 1})
    assert final["type"] == "error"
    assert final["error"]["type"] == "GurobiError"
    assert "NoSuchParameter" in final["error"]["message"]


def test_reveal_latest(api, monkeypatch):
    opened = []
    monkeypatch.setattr("subprocess.run", lambda cmd, check: opened.append(cmd))
    assert api.post("/api/runs/latest/reveal").status_code == 404
    run_id = start(api, "x = 1\n")
    events(api, run_id)
    assert api.post("/api/runs/latest/reveal").status_code == 200
    assert opened[0][:2] == ["open", "-R"] and run_id in opened[0][2]


def test_status_from_log():
    assert status_from_log(["Optimal objective  1.0"]) == "Optimal"
    assert status_from_log(["Model is infeasible"]) == "Infeasible"
    assert status_from_log(["nothing"]) is None


def test_is_same_handles_values_without_a_single_truth_value():
    class Arrayish:
        def __eq__(self, other):
            raise ValueError("ambiguous")

    assert is_same({"a": 1}, {"a": 1}) and not is_same(Arrayish(), 1)
