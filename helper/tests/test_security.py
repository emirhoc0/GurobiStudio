import stat

import pytest
from fastapi.testclient import TestClient

from gurobistudio.main import create_app


@pytest.fixture
def env(tmp_path):
    dist = tmp_path / "dist"
    dist.mkdir()
    (dist / "taskpane.html").write_text('<meta name="gs-token" content="__GS_TOKEN__">')
    (dist / "a.js").write_text("1")
    data = tmp_path / "data"
    client = TestClient(create_app(data_dir=data, dist_dir=dist), base_url="https://localhost:8765")
    token = (data / "token").read_text()
    client.page_url = f"/taskpane.html?k={(data / 'page_key').read_text()}"
    return client, token, data


def auth(token):
    return {"Authorization": f"Bearer {token}"}


def test_token_file_is_private_and_stable(env, tmp_path):
    _, token, data = env
    assert stat.S_IMODE((data / "token").stat().st_mode) == 0o600
    assert stat.S_IMODE((data / "page_key").stat().st_mode) == 0o600
    create_app(data_dir=data, dist_dir=tmp_path / "dist")
    assert (data / "token").read_text() == token


def test_api_requires_token(env):
    client, token, _ = env
    assert client.get("/api/health").status_code == 401
    assert client.get("/api/health", headers=auth("wrong")).status_code == 401
    assert client.get("/api/health", headers=auth(token)).status_code == 200


def test_bad_host_rejected_even_with_token(env):
    client, token, _ = env
    r = client.get("/api/health", headers={**auth(token), "Host": "evil.example:8765"})
    assert r.status_code == 403
    assert client.get("/taskpane.html", headers={"Host": "evil.example:8765"}).status_code == 403


def test_wrong_port_host_rejected(env):
    client, _, _ = env
    assert client.get("/", headers={"Host": "localhost:9999"}).status_code == 403


def test_bad_origin_rejected_even_with_token(env):
    client, token, _ = env
    for origin in ("https://evil.example", "null", "http://localhost:8765"):
        r = client.get("/api/health", headers={**auth(token), "Origin": origin})
        assert r.status_code == 403, origin


def test_same_origin_allowed(env):
    client, token, _ = env
    r = client.get("/api/health", headers={**auth(token), "Origin": "https://localhost:8765"})
    assert r.status_code == 200


def test_no_cors_headers_and_corp_set(env):
    client, token, _ = env
    r = client.get("/api/health", headers={**auth(token), "Origin": "https://localhost:8765"})
    assert "access-control-allow-origin" not in r.headers
    assert r.headers["cross-origin-resource-policy"] == "same-origin"


def test_page_gets_token_and_static_does_not_need_it(env):
    client, token, _ = env
    r = client.get(client.page_url)
    assert token in r.text and "__GS_TOKEN__" not in r.text
    assert r.headers["cache-control"] == "no-store"
    assert client.get("/a.js").status_code == 200


def test_health_shape(env):
    client, token, _ = env
    body = client.get("/api/health", headers=auth(token)).json()
    assert body["gurobipy"] == "13.0.3"
    assert "quicksum" in body["reserved_names"]
    assert body["python"].startswith("3.13")


def test_page_withholds_token_without_key(env):
    client, token, _ = env
    for url in ("/", "/taskpane.html", "/taskpane.html?k=wrong", "/?k="):
        r = client.get(url)
        assert r.status_code == 403, url
        assert token not in r.text
