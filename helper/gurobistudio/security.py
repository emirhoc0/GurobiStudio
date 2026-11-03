"""Host/Origin/token checks. Models run arbitrary Python as the user, so these guard the API."""

import os
import secrets
import sys
import tempfile
from pathlib import Path

from starlette.responses import JSONResponse


def load_or_create_secret(data_dir: Path, name: str) -> str:
    """Create atomically so a crash or a concurrent start can never leave an empty secret."""
    data_dir.mkdir(parents=True, exist_ok=True)
    path = data_dir / name
    if not path.exists():
        fd, tmp = tempfile.mkstemp(dir=data_dir)  # mkstemp creates the file as 0600
        try:
            with os.fdopen(fd, "w") as f:
                f.write(secrets.token_urlsafe(32))
            os.link(tmp, path)
        except FileExistsError:
            pass  # another process won; use its value
        finally:
            os.unlink(tmp)
    secret = path.read_text().strip()
    if not secret:
        raise RuntimeError(f"{path} is empty. Delete it and re-run scripts/install.sh.")
    return secret


class SecurityMiddleware:
    """Plain ASGI so SSE responses are never buffered."""

    def __init__(self, app, port: int, token: str):
        self.app = app
        self.hosts = {f"localhost:{port}".encode(), f"127.0.0.1:{port}".encode()}
        self.origins = {b"https://" + h for h in self.hosts}
        self.auth = f"Bearer {token}".encode()

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http":
            return await self.app(scope, receive, send)

        headers = dict(scope["headers"])
        problem = self._check(scope["path"], headers)
        if problem:
            print(f"rejected {scope['path']}: {problem}", file=sys.stderr)
            status = 401 if problem == "bad token" else 403
            return await JSONResponse({"error": problem}, status)(scope, receive, send)

        async def send_with_headers(message):
            if message["type"] == "http.response.start":
                message["headers"] = [
                    *message.get("headers", []),
                    (b"cross-origin-resource-policy", b"same-origin"),
                ]
            await send(message)

        await self.app(scope, receive, send_with_headers)

    def _check(self, path: str, headers: dict[bytes, bytes]) -> str | None:
        if headers.get(b"host") not in self.hosts:
            return "bad host"  # blocks DNS rebinding
        origin = headers.get(b"origin")
        if origin is not None and origin not in self.origins:
            return f"bad origin {origin.decode(errors='replace')}"
        if path.startswith("/api/"):
            supplied = headers.get(b"authorization", b"")
            if not secrets.compare_digest(supplied, self.auth):
                return "bad token"
        return None
