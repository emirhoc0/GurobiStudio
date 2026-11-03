import asyncio
import functools
import json
import os
import secrets
import sys
import time
from importlib import metadata
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

from . import __version__
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
