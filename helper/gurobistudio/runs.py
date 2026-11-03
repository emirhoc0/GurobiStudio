"""Starts, streams and stops solves. One at a time per helper."""

import hashlib
import json
import os
import re
import secrets
import shutil
import signal
import subprocess
import sys
import threading
import time
from pathlib import Path

# Gurobi logs these parameter names; drop the lines so nothing about the WLS login can leak.
CREDENTIAL_LINE = re.compile(r"wls\s*(secret|accessid)|licenseid", re.IGNORECASE)
LOG_STATUS = (
    (re.compile(r"optimal (objective|solution)", re.IGNORECASE), "Optimal"),
    (re.compile(r"(model is )?infeasible( model)?\b", re.IGNORECASE), "Infeasible"),
    (re.compile(r"time limit reached", re.IGNORECASE), "Time limit"),
)
HELPER_ROOT = Path(__file__).resolve().parents[1]
KILL_AFTER_S = 3


def status_from_log(lines: list[str]) -> str | None:
    # ponytail: only for models created inside functions; matches Gurobi's English log text.
    for line in reversed(lines):
        for pattern, name in LOG_STATUS:
            if pattern.search(line):
                return name
    return None


class Run:
    def __init__(self, run_id: str, run_dir: Path, process: subprocess.Popen):
        self.id = run_id
        self.dir = run_dir
        self.process = process
        self.events: list[dict] = [{"type": "status", "state": "running"}]
        self.finished = False
        self.cancelled = False
        self.started = time.monotonic()


class RunManager:
    def __init__(self, runs_dir: Path, keep: int = 10):
        self.runs_dir = runs_dir
        self.keep = keep
        self.runs: dict[str, Run] = {}
        self.lock = threading.Lock()

    def busy(self) -> bool:
        return any(not r.finished for r in self.runs.values())

    def start(self, spec: dict) -> Run | None:
        with self.lock:
            if self.busy():
                return None
            run_id = f"{time.time_ns()}-{secrets.token_hex(3)}"
            run_dir = self.runs_dir / run_id
            run_dir.mkdir(parents=True)
            (run_dir / "input.json").write_text(json.dumps(spec))
            process = subprocess.Popen(
                [sys.executable, "-u", "-m", "gurobistudio.runner", str(run_dir)],
                cwd=HELPER_ROOT,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                stdin=subprocess.DEVNULL,
                text=True,
                errors="replace",
                bufsize=1,
                start_new_session=True,  # its own process group, so Stop reaches everything
            )
            run = Run(run_id, run_dir, process)
            self.runs[run_id] = run
        threading.Thread(target=self._pump, args=(run,), daemon=True).start()
        self._prune()
        return run

    def _pump(self, run: Run):
        lines: list[str] = []
        with open(run.dir / "log.txt", "w") as log:
            for raw in run.process.stdout:
                line = raw.rstrip("\n")
                if CREDENTIAL_LINE.search(line):
                    continue
                lines.append(line)
                log.write(line + "\n")
                log.flush()
                run.events.append({"type": "log", "line": line})
        run.process.wait()
        run.events.extend(self._final_events(run, lines))
        run.finished = True

    def _final_events(self, run: Run, lines: list[str]) -> list[dict]:
        if run.cancelled:
            return [{"type": "status", "state": "stopped"}]
        elapsed = round(time.monotonic() - run.started, 2)
        try:
            result = json.loads((run.dir / "result.json").read_text())
        except (OSError, ValueError):
            message = f"The model process ended unexpectedly (exit code {run.process.returncode})."
            result = {"ok": False, "error": {"type": "Crash", "message": message, "line": None}}
        if not result["ok"]:
            return [{"type": "error", "elapsed": elapsed, "error": result["error"]}]
        result["status"] = result["status"] or status_from_log(lines) or "Done"
        return [{"type": "result", "elapsed": elapsed, **result}]

    def cancel(self, run: Run):
        if run.finished:
            return
        run.cancelled = True
        pgid = run.process.pid  # start_new_session makes the child its group leader
        self._signal(pgid, signal.SIGINT)  # lets Gurobi stop and release its license token
        timer = threading.Timer(KILL_AFTER_S, self._signal, (pgid, signal.SIGKILL))
        timer.daemon = True
        timer.start()

    @staticmethod
    def _signal(pgid: int, sig: signal.Signals):
        try:
            os.killpg(pgid, sig)
        except ProcessLookupError:
            pass

    def _prune(self):
        dirs = sorted(
            (p for p in self.runs_dir.iterdir() if p.is_dir()),
            key=lambda p: int(p.name.split("-")[0]),
        )
        for old in dirs[: -self.keep]:
            shutil.rmtree(old, ignore_errors=True)
            self.runs.pop(old.name, None)


class TrustStore:
    """Approved models live here, not in the workbook, so a workbook can't approve itself."""

    def __init__(self, path: Path):
        self.path = path

    def _load(self) -> set[str]:
        try:
            return set(json.loads(self.path.read_text()))
        except (OSError, ValueError):
            return set()

    @staticmethod
    def fingerprint(model: str, workbook_id: str) -> str:
        return hashlib.sha256(f"{workbook_id}\0{model}".encode()).hexdigest()

    def is_trusted(self, fingerprint: str) -> bool:
        return fingerprint in self._load()

    def add(self, fingerprint: str):
        self.path.write_text(json.dumps(sorted(self._load() | {fingerprint})))
