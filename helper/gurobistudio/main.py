import asyncio
import functools
import json
import os
import secrets
import subprocess
import sys
import time
from importlib import metadata
from pathlib import Path

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

from . import __version__, settings
from .license import check_license
from .runs import Run, RunManager, TrustStore
from .security import SecurityMiddleware, load_or_create_secret

REPO = Path(__file__).resolve().parents[2]
DEFAULT_PORT = 8765


def default_data_dir() -> Path:
    env = os.environ.get("GS_DATA_DIR")
    return Path(env) if env else Path.home() / "Library/Application Support/GurobiStudio"


@functools.cache
def gurobi_info() -> dict:
    try:
        import gurobipy  # importing does not touch the license

        return {"gurobipy": metadata.version("gurobipy"), "reserved_names": list(gurobipy.__all__)}
    except ImportError:
        return {"gurobipy": None, "reserved_names": []}


def create_app(
    port: int = DEFAULT_PORT,
    data_dir: Path | None = None,
    dist_dir: Path = REPO / "add-in" / "dist",
) -> FastAPI:
    data_dir = data_dir or default_data_dir()
    token = load_or_create_secret(data_dir, "token")
    page_key = load_or_create_secret(data_dir, "page_key")
    app = FastAPI(docs_url=None, redoc_url=None, openapi_url=None)
    app.add_middleware(SecurityMiddleware, port=port, token=token)

    @app.get("/api/health")
    def health():
        return {
            "version": __version__,
            "python": sys.version.split()[0],
            **gurobi_info(),
        }

    settings_path = data_dir / "settings.json"
    runs = RunManager(data_dir / "runs", settings.load(settings_path))
    trust = TrustStore(data_dir / "trusted_models.json")

    def get_run(run_id: str) -> Run:
        run = runs.runs.get(run_id)
        if run is None:
            raise HTTPException(404, "Unknown run")
        return run

    @app.post("/api/run")
    async def start_run(request: Request):
        spec = await request.json()
        fingerprint = trust.fingerprint(spec["model"], spec["workbook_id"])
        if not trust.is_trusted(fingerprint):
            if not spec.get("approved"):
                return JSONResponse({"trust_required": True}, 409)
            trust.add(fingerprint)
        run = runs.start({k: spec.get(k) for k in ("model", "items", "params")})
        if run is None:
            return JSONResponse({"busy": True}, 409)
        return {"run_id": run.id}

    @app.get("/api/run/{run_id}/events")
    async def run_events(run_id: str):
        run = get_run(run_id)

        async def stream():
            sent = 0
            while True:
                finished = run.finished  # read first, so nothing added before it is missed
                while sent < len(run.events):
                    yield f"data: {json.dumps(run.events[sent])}\n\n"
                    sent += 1
                if finished:
                    return
                await asyncio.sleep(0.05)  # ponytail: polling; one local client, so cost is nil

        return StreamingResponse(stream(), media_type="text/event-stream")

    @app.post("/api/run/{run_id}/cancel")
    def cancel_run(run_id: str):
        runs.cancel(get_run(run_id))
        return {"ok": True}

    @app.get("/api/run/{run_id}/files")
    def run_files(run_id: str):
        run = get_run(run_id)
        files = [{"name": p.name, "size": p.stat().st_size} for p in sorted(run.dir.iterdir())]
        return {"dir": str(run.dir), "files": files}

    def reveal(run_dir: Path):
        subprocess.run(["open", "-R", str(run_dir / "model.py")], check=False)
        return {"ok": True}

    @app.post("/api/run/{run_id}/reveal")
    def reveal_run(run_id: str):
        return reveal(get_run(run_id).dir)

    @app.post("/api/runs/latest/reveal")
    def reveal_latest():
        latest = runs.latest_dir()
        if latest is None:
            raise HTTPException(404, "No runs yet")
        return reveal(latest)

    @app.get("/api/settings")
    def get_settings():
        return {**runs.settings, "port": port, "default_python": sys.executable}

    @app.put("/api/settings")
    async def put_settings(request: Request):
        try:
            cleaned = settings.clean(await request.json())
        except ValueError as e:
            return JSONResponse({"error": str(e)}, 400)
        settings_path.write_text(json.dumps(cleaned))
        runs.settings = cleaned
        return get_settings()

    @app.post("/api/license/check")
    def license_check():
        if runs.busy():
            return JSONResponse({"busy": True}, 409)  # a solve may hold the only WLS session
        return check_license(runs.python(), settings.child_env(runs.settings))

    @app.post("/api/diag")
    async def diag(request: Request):
        report = await request.json()
        report["server_saw"] = {
            "origin": request.headers.get("origin"),
            "host": request.headers.get("host"),
        }
        diag_dir = data_dir / "diag"
        diag_dir.mkdir(exist_ok=True)
        text = json.dumps(report, indent=2)
        (diag_dir / f"{int(time.time())}.json").write_text(text)
        (diag_dir / "latest.json").write_text(text)
        return {"saved": str(diag_dir / "latest.json")}

    @app.get("/api/diag/stream")
    async def diag_stream():
        # Proves fetch() in Excel's web view delivers SSE incrementally.
        async def events():
            for i in range(4):
                yield f"data: {json.dumps({'i': i})}\n\n"
                await asyncio.sleep(0.4)

        return StreamingResponse(events(), media_type="text/event-stream")

    def page(k: str = ""):
        # The API token is only handed to holders of the key in the private sideloaded manifest,
        # so other local programs and accounts can't read it from this page.
        if not secrets.compare_digest(k.encode(), page_key.encode()):
            return JSONResponse({"error": "missing page key"}, 403)
        html = (dist_dir / "taskpane.html").read_text().replace("__GS_TOKEN__", token)
        return HTMLResponse(html, headers={"Cache-Control": "no-store"})

    app.get("/", include_in_schema=False)(page)
    app.get("/taskpane.html", include_in_schema=False)(page)
    if dist_dir.exists():
        app.mount("/", StaticFiles(directory=dist_dir), name="static")
    return app
