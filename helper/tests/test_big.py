"""A million cells through the real API. No Gurobi, so this measures our code alone."""

import time
from pathlib import Path

from test_runs import api, changed, events, item, solve, start  # noqa: F401

from gurobistudio import runs as runs_module
from gurobistudio.runner import is_same

CELLS = 1_000_000


def test_a_million_cells_arrive_intact(api):  # noqa: F811
    side = int(CELLS**0.5)
    pairs = [[[f"p{r}", c], r * c + 0.5] for r in range(side) for c in range(side)]
    model = "out = {k: v * 2 for k, v in cost.items()}\ncheck = len(cost)\n"
    started = time.monotonic()
    final, _ = solve(
        api,
        model,
        [
            item("cost", {"__dict__": pairs}, readOnly=True),
            item("out", {"__dict__": []}),
            item("check", 0),
        ],
    )
    assert final["type"] == "result"
    result = changed(final)
    assert result["check"] == CELLS
    assert result["out"]["__dict__"][-1] == [["p999", 999], (999 * 999 + 0.5) * 2]
    assert len(result["out"]["__dict__"]) == CELLS
    assert time.monotonic() - started < 60  # about 8 s when written; a big slowdown means a bug


def test_in_place_edit_of_a_large_item_is_seen(api):  # noqa: F811
    # The snapshot is a one-level copy, so this fails if it ever shares the item's container.
    values = list(range(100_000))
    final, _ = solve(api, "data[99_999] = -1\n", [item("data", values)])
    assert changed(final)["data"][-1] == -1


def test_snapshot_comparison_sees_a_replaced_tuple():
    assert not is_same([(1, 2), (3, 9)], [(1, 2), (3, 4)])


def test_chatty_model_keeps_memory_bounded_and_the_result(api, monkeypatch):  # noqa: F811
    monkeypatch.setattr(runs_module, "MAX_BUFFERED_EVENTS", 1000)
    run_id = start(api, "for i in range(5000):\n    print(i)\nout = 1\n", [item("out", 0)])
    run_dir = Path(api.get(f"/api/run/{run_id}/files").json()["dir"])
    deadline = time.time() + 20
    while not (run_dir / "result.json").exists() and time.time() < deadline:
        time.sleep(0.02)
    time.sleep(0.3)  # the pump finishes its last lines
    evs = events(api, run_id)
    assert len(evs) <= 1000
    assert evs[-1]["type"] == "result" and changed(evs[-1]) == {"out": 1}
    assert len((run_dir / "log.txt").read_text().splitlines()) == 5001  # every line, plus the summary


def test_events_are_released_once_sent(api):  # noqa: F811
    run_id = start(api, "x = 1\n")
    assert events(api, run_id)[-1]["type"] == "result"
    assert events(api, run_id) == []
