"""Pytest fixtures for contract-importer-backend."""
import io
import os
import sys
from pathlib import Path

import pytest

# Ensure repo root is on sys.path so `import app` works
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch):
    """Reset sensitive env vars before each test."""
    for key in ("OPENAI_API_KEY", "GOOGLE_CREDENTIALS_JSON", "CORS_ORIGINS",
                "OPENAI_MODEL", "OPENAI_MAX_TOKENS", "MAX_UPLOAD_MB",
                "SPREADSHEET_ID", "SHEET_GID"):
        monkeypatch.delenv(key, raising=False)
    yield


@pytest.fixture
def client():
    # Reimport app so monkeypatches applied before fixture take effect
    import importlib
    import app as app_module
    importlib.reload(app_module)
    flask_app = app_module.app
    flask_app.config["TESTING"] = True
    with flask_app.test_client() as c:
        yield c


@pytest.fixture
def tiny_png_bytes():
    """1x1 PNG, valid minimal image."""
    import base64
    # 1x1 transparent PNG
    return base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgAAIAAAUAAen63NgAAAAASUVORK5CYII="
    )
