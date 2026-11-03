"""Real checks use this machine's license; the parser is pinned on sample logs."""

import pytest
from fastapi.testclient import TestClient

from gurobistudio.license import parse_license
from gurobistudio.main import create_app

BASE = "https://localhost:8765"


@pytest.fixture
def api(tmp_path):
    dist = tmp_path / "dist"
    dist.mkdir()
    (dist / "taskpane.html").write_text("x")
    client = TestClient(create_app(data_dir=tmp_path / "data", dist_dir=dist), base_url=BASE)
    client.headers["Authorization"] = f"Bearer {(tmp_path / 'data' / 'token').read_text()}"
    return client


def test_academic_wls():
    log = "Set parameter LicenseID to value 1\nAcademic license 1 - registered to a@b.c\n"
    out = parse_license(log, {"expiry": 20270115, "size_ok": True})
    assert out["ok"] and out["type"] == "Academic" and not out["restricted"]
    assert out["expiry"] == "2027-01-15"
    assert "a@b.c" not in str(out)


def test_restricted_by_banner():
    log = "Restricted license - for non-production use only - expires 2027-11-29\n"
    out = parse_license(log, {"expiry": 0, "size_ok": False})
    assert out["restricted"] and out["type"] == "Restricted"
    assert out["expiry"] == "2027-11-29"
    assert "size-limited" in out["message"]


def test_size_error_alone_means_restricted():
    # The wording of the banner can change; the size test can't lie.
    assert parse_license("Some new banner text", {"size_ok": False})["restricted"]


def test_no_expiry_values_are_hidden():
    assert parse_license("", {"expiry": 99991231, "size_ok": True})["expiry"] is None


@pytest.mark.parametrize(
    "error",
    ["Could not connect to token server", "Timed out contacting WLS", "Connection refused"],
)
def test_offline_is_said_plainly(error):
    out = parse_license("", {"error": error})
    assert not out["ok"] and out["offline"] and "internet" in out["message"]


def test_other_errors_keep_their_message_without_secrets():
    assert parse_license("", {"error": "License expired"})["message"] == "License expired"
    leaked = parse_license("", {"error": "bad WLSSecret abc123"})["message"]
    assert "abc123" not in leaked


def test_real_license(api):
    out = api.post("/api/license/check").json()
    assert out["ok"], out
    assert out["type"] == "Academic"
    assert not out["restricted"]


def test_settings_round_trip_and_validation(api, tmp_path):
    defaults = api.get("/api/settings").json()
    assert defaults["keep_runs"] == 10 and defaults["port"] == 8765
    assert api.put("/api/settings", json={"keep_runs": 3}).json()["keep_runs"] == 3
    assert api.get("/api/settings").json()["keep_runs"] == 3
    bad = api.put("/api/settings", json={"python_path": str(tmp_path / "nope")})
    assert bad.status_code == 400 and "not an executable" in bad.json()["error"]
    assert api.put("/api/settings", json={"keep_runs": 0}).status_code == 400
    assert api.get("/api/settings").json()["keep_runs"] == 3


def test_license_path_reaches_the_run(api, tmp_path):
    lic = tmp_path / "x.lic"
    lic.write_text("")
    api.put("/api/settings", json={"license_path": str(lic)})
    from gurobistudio import settings

    assert settings.child_env(settings.clean({"license_path": str(lic)}))["GRB_LICENSE_FILE"]


def test_retention_setting_prunes(api):
    api.put("/api/settings", json={"keep_runs": 2})
    last = None
    for _ in range(4):
        last = api.post(
            "/api/run",
            json={"model": "x = 1\n", "items": [], "params": {}, "workbook_id": "w", "approved": True},
        ).json()["run_id"]
        api.get(f"/api/run/{last}/events")
    folder = api.get(f"/api/run/{last}/files").json()["dir"]
    from pathlib import Path

    assert len([p for p in Path(folder).parent.iterdir() if p.is_dir()]) == 2
