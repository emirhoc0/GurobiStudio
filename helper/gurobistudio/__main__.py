import os
from pathlib import Path

import uvicorn

from .main import DEFAULT_PORT, create_app, default_data_dir

# install.sh records the port it wrote into the manifest, so the two can't disagree.
port_file = default_data_dir() / "port"
saved = port_file.read_text().strip() if port_file.exists() else DEFAULT_PORT
port = int(os.environ.get("GS_PORT", saved))
certs = Path.home() / ".office-addin-dev-certs"
uvicorn.run(
    create_app(port=port),
    host="127.0.0.1",
    port=port,
    ssl_certfile=certs / "localhost.crt",
    ssl_keyfile=certs / "localhost.key",
    log_level="warning",
)
