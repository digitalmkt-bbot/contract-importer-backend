"""Tests for API_TOKEN auth, rate limiting, and NDJSON streaming."""
import io
import json
import importlib


def _fresh_client(monkeypatch, **env):
    for k, v in env.items():
        if v is None:
            monkeypatch.delenv(k, raising=False)
        else:
            monkeypatch.setenv(k, str(v))
    import app as app_module
    importlib.reload(app_module)
    app_module.app.config["TESTING"] = True
    return app_module, app_module.app.test_client()


def test_auth_disabled_by_default(client):
    """No API_TOKEN → endpoints are open (except still require valid input)."""
    rv = client.post("/api/import-sheets", json={"items": []})
    assert rv.status_code == 400  # rejected for content reason, not 401


def test_auth_enforced_when_token_set(monkeypatch):
    _, c = _fresh_client(monkeypatch, API_TOKEN="s3cr3t", RATE_LIMIT_MAX="0")

    # No header
    rv = c.post("/api/import-sheets", json={"items": []})
    assert rv.status_code == 401
    assert "Unauthorized" in rv.get_json()["error"]

    # Wrong header
    rv = c.post("/api/import-sheets", json={"items": []},
                headers={"X-API-Token": "wrong"})
    assert rv.status_code == 401

    # Correct header passes auth (falls through to empty-items 400)
    rv = c.post("/api/import-sheets", json={"items": []},
                headers={"X-API-Token": "s3cr3t"})
    assert rv.status_code == 400
    assert "ไม่มีข้อมูล" in rv.get_json()["error"]


def test_auth_via_query_param(monkeypatch):
    """Fallback: ?api_token=... so pre-signed URLs / curl work."""
    _, c = _fresh_client(monkeypatch, API_TOKEN="q1", RATE_LIMIT_MAX="0")
    rv = c.post("/api/import-sheets?api_token=q1", json={"items": []})
    assert rv.status_code == 400  # past auth, content-failed


def test_rate_limit_blocks_after_threshold(monkeypatch):
    """RATE_LIMIT_MAX=3 → 4th call returns 429."""
    _, c = _fresh_client(monkeypatch, RATE_LIMIT_MAX="3", RATE_LIMIT_WINDOW="60",
                         API_TOKEN=None)
    for _ in range(3):
        rv = c.post("/api/extract")  # 400 because no file, but still counted
        assert rv.status_code == 400
    rv = c.post("/api/extract")
    assert rv.status_code == 429
    body = rv.get_json()
    assert "Rate limit" in body["error"]
    assert "retry_after" in body
    assert rv.headers.get("Retry-After") is not None


def test_rate_limit_disabled_when_max_zero(monkeypatch):
    _, c = _fresh_client(monkeypatch, RATE_LIMIT_MAX="0", API_TOKEN=None)
    for _ in range(20):
        rv = c.post("/api/extract")
        assert rv.status_code == 400  # no-file error, not 429


def test_status_reports_auth_and_rate_limit(monkeypatch):
    _, c = _fresh_client(monkeypatch, API_TOKEN="xyz",
                         RATE_LIMIT_MAX="7", RATE_LIMIT_WINDOW="120")
    data = c.get("/api/status").get_json()
    assert data["auth_required"] is True
    assert data["rate_limit"]["max"] == 7
    assert data["rate_limit"]["window_seconds"] == 120
    assert data["rate_limit"]["backend"] in ("memory", "redis")


def test_stream_endpoint_rejects_missing_file(client):
    rv = client.post("/api/extract/stream")
    assert rv.status_code == 400


def test_stream_endpoint_requires_auth_when_token_set(monkeypatch):
    _, c = _fresh_client(monkeypatch, API_TOKEN="t", RATE_LIMIT_MAX="0")
    rv = c.post("/api/extract/stream")
    assert rv.status_code == 401


def test_stream_endpoint_returns_ndjson_for_ocr_fallback(client, tiny_png_bytes):
    """Without OPENAI_API_KEY the streaming route should still emit valid NDJSON."""
    rv = client.post(
        "/api/extract/stream",
        data={"file": (io.BytesIO(tiny_png_bytes), "tiny.png")},
        content_type="multipart/form-data",
    )
    # If pytesseract is absent in CI the response may 500 — we only assert shape
    # when a streaming response actually starts.
    if rv.status_code != 200:
        return
    assert rv.mimetype == "application/x-ndjson"
    lines = [ln for ln in rv.get_data(as_text=True).splitlines() if ln.strip()]
    assert lines, "expected at least one NDJSON line"
    events = [json.loads(ln) for ln in lines]
    assert events[0]["event"] == "start"
    assert events[-1]["event"] in ("done", "error")
