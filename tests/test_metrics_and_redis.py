"""Tests for /metrics endpoint and Redis rate-limit backend fallback."""
import importlib
import json
import pytest


def _fresh(monkeypatch, **env):
    import sys
    for k, v in env.items():
        monkeypatch.setenv(k, v)
    if "app" in sys.modules:
        del sys.modules["app"]
    import app as app_mod
    app_mod.app.config["TESTING"] = True
    return app_mod, app_mod.app.test_client()


def test_metrics_endpoint_json_default(monkeypatch):
    mod, c = _fresh(monkeypatch)
    c.get("/api/status")
    c.get("/api/status")
    r = c.get("/metrics")
    assert r.status_code == 200
    assert r.is_json
    data = r.get_json()
    assert "requests_total" in data
    assert data["requests_total"] >= 2
    assert "requests_by_route" in data


def test_metrics_endpoint_prom_format(monkeypatch):
    mod, c = _fresh(monkeypatch)
    c.get("/api/status")
    r = c.get("/metrics?format=prom")
    assert r.status_code == 200
    text = r.get_data(as_text=True)
    assert "# HELP" in text
    assert "contract_importer_requests_total" in text


def test_metrics_endpoint_prom_via_accept_header(monkeypatch):
    mod, c = _fresh(monkeypatch)
    c.get("/api/status")
    r = c.get("/metrics", headers={"Accept": "text/plain"})
    assert r.status_code == 200
    text = r.get_data(as_text=True)
    assert "contract_importer_" in text


def test_metrics_endpoint_requires_auth_when_token_set(monkeypatch):
    mod, c = _fresh(monkeypatch, API_TOKEN="secret")
    r = c.get("/metrics")
    assert r.status_code == 401
    r2 = c.get("/metrics", headers={"X-API-Token": "secret"})
    assert r2.status_code == 200


def test_status_reports_rate_limit_backend_memory(monkeypatch):
    mod, c = _fresh(monkeypatch, RATE_LIMIT_MAX="5")
    data = c.get("/api/status").get_json()
    assert data["rate_limit"]["backend"] == "memory"


def test_status_reports_metrics_enabled(monkeypatch):
    mod, c = _fresh(monkeypatch)
    data = c.get("/api/status").get_json()
    assert data.get("metrics_enabled") is True


def test_redis_client_returns_none_when_redis_url_unset(monkeypatch):
    mod, _ = _fresh(monkeypatch)
    # _redis_client should return None when REDIS_URL is not set
    mod._redis_singleton = None  # reset singleton if present
    client = mod._redis_client()
    assert client is None


def test_redis_client_returns_none_on_import_or_conn_error(monkeypatch):
    mod, _ = _fresh(monkeypatch, REDIS_URL="redis://invalid-host:6379/0")
    # redis package may not be installed, or connection will fail. Either way: None.
    mod._redis_singleton = None
    client = mod._redis_client()
    assert client is None


def test_rate_limit_check_falls_back_to_memory_when_redis_unavailable(monkeypatch):
    mod, c = _fresh(monkeypatch, RATE_LIMIT_MAX="2", RATE_LIMIT_WINDOW="60",
                    REDIS_URL="redis://invalid-host:1/0")
    # /api/extract is protected by require_auth (which applies rate-limit).
    # No file -> 400, but the 3rd request should be rejected by the limiter (429).
    assert c.post("/api/extract").status_code == 400
    assert c.post("/api/extract").status_code == 400
    r = c.post("/api/extract")
    assert r.status_code == 429
    assert r.headers.get("Retry-After")
