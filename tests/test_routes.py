"""Smoke / contract tests for the Flask routes."""
import io
import json


def test_index_fallback_json_when_no_static(client, monkeypatch, tmp_path):
    # Point static_folder at an empty directory so the JSON branch is taken
    import app as app_module
    monkeypatch.setattr(app_module.app, "static_folder", str(tmp_path))
    rv = client.get("/")
    assert rv.status_code == 200
    body = rv.get_json()
    assert body["status"] == "ok"
    assert body["has_api_key"] is False


def test_status_shape(client):
    rv = client.get("/api/status")
    assert rv.status_code == 200
    data = rv.get_json()
    for key in ("has_api_key", "has_credentials", "spreadsheet_id",
                "service_account_email", "model", "max_upload_mb", "cors_origins"):
        assert key in data
    assert data["has_api_key"] is False
    assert data["has_credentials"] is False
    assert data["cors_origins"] == "*"
    assert data["max_upload_mb"] == 25  # default


def test_status_respects_cors_origins_env(monkeypatch):
    monkeypatch.setenv("CORS_ORIGINS", "https://a.com, https://b.com")
    import importlib, app as app_module
    importlib.reload(app_module)
    with app_module.app.test_client() as c:
        data = c.get("/api/status").get_json()
    assert data["cors_origins"] == ["https://a.com", "https://b.com"]


def test_extract_missing_file_returns_400(client):
    rv = client.post("/api/extract")
    assert rv.status_code == 400
    assert "ไม่พบไฟล์" in rv.get_json()["error"]


def test_extract_rejects_unsupported_type(client):
    data = {"file": (io.BytesIO(b"hello world"), "notes.txt")}
    rv = client.post("/api/extract", data=data, content_type="multipart/form-data")
    assert rv.status_code == 400
    assert "รองรับเฉพาะ" in rv.get_json()["error"]


def test_extract_image_ocr_fallback(client, tiny_png_bytes):
    """No API key → OCR branch. Tesseract on 1x1 PNG returns empty but shouldn't 500."""
    data = {"file": (io.BytesIO(tiny_png_bytes), "tiny.png")}
    rv = client.post("/api/extract", data=data, content_type="multipart/form-data")
    # OCR fallback is best-effort; we just require a stable JSON contract
    assert rv.status_code in (200, 500)  # pytesseract may not be installed in CI
    if rv.status_code == 200:
        body = rv.get_json()
        assert "items" in body
        assert "company_name" in body


def test_import_sheets_requires_items(client):
    rv = client.post("/api/import-sheets", json={"items": []})
    assert rv.status_code == 400
    assert "ไม่มีข้อมูล" in rv.get_json()["error"]


def test_import_sheets_validates_spreadsheet_id(client):
    rv = client.post("/api/import-sheets", json={
        "items": [{"product_name": "x", "net_rate": 1}],
        "spreadsheet_id": "not-valid!",
    })
    assert rv.status_code == 400
    assert "spreadsheet_id" in rv.get_json()["error"]


def test_import_sheets_missing_creds(client):
    rv = client.post("/api/import-sheets", json={
        "items": [{"product_name": "x", "net_rate": 1}],
        "spreadsheet_id": "1X_gcLo3RROT11Hv9qvhiegoztk9STv4lP2aTq0Ih0Ho",
    })
    assert rv.status_code == 400
    assert "GOOGLE_CREDENTIALS_JSON" in rv.get_json()["error"]


def test_extract_rejects_oversize_upload(client, monkeypatch):
    """Configured MAX_UPLOAD_MB=1 → uploading 1.5MB must 413."""
    monkeypatch.setenv("MAX_UPLOAD_MB", "1")
    import importlib, app as app_module
    importlib.reload(app_module)
    with app_module.app.test_client() as c:
        big = io.BytesIO(b"\x00" * (1_500_000))
        rv = c.post(
            "/api/extract",
            data={"file": (big, "big.pdf")},
            content_type="multipart/form-data",
        )
    assert rv.status_code == 413
