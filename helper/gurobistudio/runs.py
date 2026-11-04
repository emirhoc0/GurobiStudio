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
from collections import deque
from pathlib import Path

from .settings import DEFAULTS, child_env

# Gurobi logs these parameter names; drop the lines so nothing about the WLS login can leak.
CREDENTIAL_LINE = re.compile(r"wls[\s_-]*(secret|access[\s_-]*id)|license[\s_-]*id", re.IGNORECASE)
# Gurobi's academic banner names the account's email. Copy log makes that easy to paste by accident.
REGISTERED_TO = re.compile(r"(registered to\s+)\S+", re.IGNORECASE)
LOG_STATUS = (
    (re.compile(r"optimal (objective|solution)", re.IGNORECASE), "Optimal"),
    (re.compile(r"(model is )?infeasible( model)?\b", re.IGNORECASE), "Infeasible"),
    (re.compile(r"time limit reached", re.IGNORECASE), "Time limit"),
)
HELPER_ROOT = Path(__file__).resolve().parents[1]
KILL_AFTER_S = 3
# Lines waiting for the pane. A chatty model drops the oldest; the full log is in the run folder.
MAX_BUFFERED_EVENTS = 50_000
STATUS_LOOKBACK_LINES = 200


def status_from_log(lines) -> str | None:
    # ponytail: only for models created inside functions; matches Gurobi's English log text.
    for line in reversed(lines):
        for pattern, name in LOG_STATUS:
            if pattern.search(line):
                return name
    return None


def scrub_error(error: dict) -> dict:
    """A traceback quotes source lines and messages, so it can hold what the log filter removes."""
    clean = dict(error)
    if CREDENTIAL_LINE.search(str(error.get("message"))):
        clean["message"] = "(hidden, because it names a license credential)"
    if error.get("traceback"):
        kept = [x for x in error["traceback"].split("\n") if not CREDENTIAL_LINE.search(x)]
        clean["traceback"] = "\n".join(kept)
    return clean


class Run:
    def __init__(self, run_id: str, run_dir: Path, process: subprocess.Popen):
        self.id = run_id
        self.dir = run_dir
        self.process = process
        self.events = deque([{"type": "status", "state": "running"}], maxlen=MAX_BUFFERED_EVENTS)
        self.finished = False
        self.cancelled = False
        self.started = time.monotonic()


class RunManager:
    def __init__(self, runs_dir: Path, settings: dict | None = None):
        self.runs_dir = runs_dir
        self.settings = settings or dict(DEFAULTS)
        self.runs: dict[str, Run] = {}
        self.lock = threading.Lock()

    def python(self) -> str:
        return self.settings["python_path"] or sys.executable

    def latest_dir(self) -> Path | None:
        """The newest run folder, on disk, so it survives a helper restart."""
        dirs = self._run_dirs()
        return dirs[-1] if dirs else None

    def active(self) -> Run | None:
        return next((r for r in self.runs.values() if not r.finished), None)

    def busy(self) -> bool:
        return self.active() is not None

    def start(self, spec: dict) -> Run | None:
        with self.lock:
            if self.busy():
                return None
            run_id = f"{time.time_ns()}-{secrets.token_hex(3)}"
            run_dir = self.runs_dir / run_id
            run_dir.mkdir(parents=True)
            (run_dir / "input.json").write_text(json.dumps(spec))
            process = subprocess.Popen(
                [self.python(), "-u", "-m", "gurobistudio.runner", str(run_dir)],
                cwd=HELPER_ROOT,
                env=child_env(self.settings),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                stdin=subprocess.DEVNULL,
                encoding="utf-8",
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
        lines: deque[str] = deque(maxlen=STATUS_LOOKBACK_LINES)
        try:
            with open(run.dir / "log.txt", "w", encoding="utf-8") as log:
                for raw in run.process.stdout:
                    line = raw.rstrip("\n")
                    if CREDENTIAL_LINE.search(line):
                        continue
                    line = REGISTERED_TO.sub(r"\1(hidden)", line)
                    lines.append(line)
                    log.write(line + "\n")
                    log.flush()
                    run.events.append({"type": "log", "line": line})
            run.process.wait()
            run.events.extend(self._final_events(run, lines))
        except Exception as e:  # noqa: BLE001  # a full disk, say; never leave the helper "busy"
            self._signal(run.process.pid, signal.SIGKILL)
            message = f"The helper lost track of this run. ({e})"
            error = {"type": "Crash", "message": message, "line": None}
            run.events.append({"type": "error", "elapsed": 0, "error": error})
        finally:
            run.finished = True

    def _final_events(self, run: Run, lines) -> list[dict]:
        if run.cancelled:
            return [{"type": "status", "state": "stopped"}]
        elapsed = round(time.monotonic() - run.started, 2)
        try:
            result = json.loads((run.dir / "result.json").read_text())
        except (OSError, ValueError):
            message = f"The model process ended unexpectedly (exit code {run.process.returncode})."
            result = {"ok": False, "error": {"type": "Crash", "message": message, "line": None}}
        if not result["ok"]:
            return [{"type": "error", "elapsed": elapsed, "error": scrub_error(result["error"])}]
        result["status"] = result["status"] or status_from_log(lines) or "Done"
        return [{"type": "result", "elapsed": elapsed, **result}]

    def cancel(self, run: Run):
        if run.finished:
            return
        run.cancelled = True
        pgid = run.process.pid  # start_new_session makes the child its group leader
        self._signal(pgid, signal.SIGINT)  # lets Gurobi stop and release its license token
        timer = threading.Timer(KILL_AFTER_S, self._kill_unless_finished, (run,))
        timer.daemon = True
        timer.start()

    def _kill_unless_finished(self, run: Run):
        if not run.finished:  # afterwards the pid may belong to something else
            self._signal(run.process.pid, signal.SIGKILL)

    @staticmethod
    def _signal(pgid: int, sig: signal.Signals):
        try:
            os.killpg(pgid, sig)
        except ProcessLookupError:
            pass

    def _run_dirs(self) -> list[Path]:
        if not self.runs_dir.exists():
            return []
        return sorted(
            (p for p in self.runs_dir.iterdir() if p.is_dir() and p.name.split("-")[0].isdigit()),
            key=lambda p: int(p.name.split("-")[0]),
        )

    def _prune(self):
        for old in self._run_dirs()[: -self.settings["keep_runs"]]:
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
    def fingerprint(model: str, workbook_id: str, params: dict | None = None) -> str:
        # Text values can be a path or a server (LogFile, ComputeServer), so approval covers them.
        # Numbers can't, so changing TimeLimit doesn't prompt again.
        text = sorted((k, v) for k, v in (params or {}).items() if isinstance(v, str))
        extra = f"\0{json.dumps(text)}" if text else ""
        return hashlib.sha256(f"{workbook_id}\0{model}{extra}".encode()).hexdigest()

    def is_trusted(self, fingerprint: str) -> bool:
        return fingerprint in self._load()

    def add(self, fingerprint: str):
        self.path.write_text(json.dumps(sorted(self._load() | {fingerprint})))
