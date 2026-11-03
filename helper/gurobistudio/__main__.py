import os
from pathlib import Path

import uvicorn

from .main import DEFAULT_PORT, create_app

port = int(os.environ.get("GS_PORT", DEFAULT_PORT))
certs = Path.home() / ".office-addin-dev-certs"
uvicorn.run(
    create_app(port=port),
    host="127.0.0.1",
    port=port,
    ssl_certfile=certs / "localhost.crt",
    ssl_keyfile=certs / "localhost.key",
    log_level="warning",
)
