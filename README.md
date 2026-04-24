# Contract Importer — Backend

Flask API that extracts pricing rows from tour-operator contract PDFs (or images) and
imports them into a Google Sheet. Deployed on Railway.

**Sister repo:** [contract-importer-frontend](https://github.com/digitalmkt-bbot/contract-importer-frontend)

---

## Architecture

```
 PDF / PNG / JPG
       │
       ▼
 Flask  /api/extract  ──►  GPT-4o vision  ──►  { company, items[] }
       │                   (fallback: Tesseract OCR)
       ▼
 Flask  /api/import-sheets  ──►  gspread  ──►  Google Sheet (cols E–Q)
```

`app.py` is the entire backend. The `static/` folder ships a copy of the frontend
UI so the Railway deployment also serves the form directly at `/`.

---

## Quick start (local)

```bash
# 1) system deps for OCR / PDF rasterisation
sudo apt-get install -y poppler-utils tesseract-ocr libglib2.0-0

# 2) python deps
python -m venv .venv && source .venv/bin/activate
pip install -r requirements-dev.txt

# 3) env
cp .env.example .env      # fill in OPENAI_API_KEY + GOOGLE_CREDENTIALS_JSON

# 4) run
flask --app app run --debug --port 5000
```

The UI is available at <http://localhost:5000/> (served from `static/index.html`).

### Docker

```bash
docker build -t contract-importer .
docker run -p 8080:8080 --env-file .env contract-importer
```

---

## Environment variables

See `.env.example` for the full list. The critical ones:

| Variable                   | Required | Notes                                                  |
|----------------------------|:--------:|--------------------------------------------------------|
| `OPENAI_API_KEY`           | ✅       | GPT-4o vision extraction. No key → Tesseract fallback. |
| `GOOGLE_CREDENTIALS_JSON`  | ✅       | Full service-account JSON, one line. Share sheet w/ its email. |
| `SPREADSHEET_ID`           |          | Default sheet; requests may override via body.          |
| `SHEET_GID`                |          | Numeric worksheet id inside the spreadsheet.            |
| `CORS_ORIGINS`             |          | Comma-separated allow-list. Default `*`.                |
| `MAX_UPLOAD_MB`            |          | Request body cap. Default 25.                           |
| `OPENAI_MODEL`             |          | Default `gpt-4o-2024-08-06`.                            |

---

## API

### `GET /api/status`
Configuration snapshot — `has_api_key`, `has_credentials`, `model`,
`max_upload_mb`, `cors_origins`, `service_account_email`.

### `POST /api/extract`
Multipart form with field **`file`** (PDF / PNG / JPG / JPEG / WEBP).
Legacy field `pdf` is still accepted for backward-compatibility.

Response:
```json
{
  "company_name": "NOAH Marine",
  "items": [
    {
      "product_name": "Big Boat | Phi Phi Island Tour | Phuket → Phi Phi | Adult",
      "departure_time": "DEP. 08:00 - ARR. 17:00",
      "net_rate": 1000,
      "selling_rate": 1500,
      "notes": "Include lunch | Valid Nov 2025 - Oct 2026"
    }
  ]
}
```

### `POST /api/import-sheets`
```json
{
  "company_name": "NOAH Marine",
  "spreadsheet_id": "1X_...",
  "overwrite": false,
  "items": [ { "product_name": "...", "net_rate": 1000, ... } ]
}
```

- On first call, if rows with the same `(company | product_name + departure_time)`
  already exist, the response is **`{ "success": false, "conflict": true,
  "duplicates": [...], "new_count": N }`** — frontend must ask user, then re-POST
  with `"overwrite": true` to replace those rows in place.
- Columns written: **E** operator · **F** product+time · **G** net_rate ·
  **H** selling_rate · **Q** notes.
  Rows are written with `ws.update("E{r}:Q{r}", ...)` (not `append_rows`) so the
  sheet's table detection doesn't offset the data.

---

## Tests

```bash
pip install -r requirements-dev.txt
pytest
```

CI (`.github/workflows/ci.yml`) runs lint + tests + docker build on every push.

---

## Deployment (Railway)

`railway.toml` is already wired to build from `Dockerfile` and probe
`/api/status`. Push to `main` and Railway redeploys.

After deploying, set the env vars above in **Railway → Project → Variables**,
and share the target Google Sheet with the `client_email` printed by
`/api/status`.
