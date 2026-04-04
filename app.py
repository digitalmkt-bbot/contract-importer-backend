#!/usr/bin/env python3
"""Contract Data Importer - Backend (Railway) - Love Andaman"""

from flask import Flask, request, jsonify
from flask_cors import CORS
import os
import json
import re
import tempfile
import base64
from io import BytesIO

app = Flask(__name__)
CORS(app, origins="*")

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1KWqJVYfoaRg3DwslW2zSQmPgScPbE9Z-0v-Ijwtdpms")
SHEET_GID = int(os.environ.get("SHEET_GID", "384942453"))


@app.route("/", methods=["GET"])
def health():
    return jsonify({
        "status": "ok",
        "service": "Contract Importer API - Love Andaman",
        "has_api_key": bool(os.environ.get("OPENAI_API_KEY")),
        "has_credentials": bool(os.environ.get("GOOGLE_CREDENTIALS_JSON"))
    })


@app.route("/api/extract", methods=["POST"])
def extract():
    """Accept PDF file upload, convert to images, call GPT-4o vision."""
    import openai
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file uploaded"}), 400

        file = request.files["file"]
        if not file.filename.lower().endswith(".pdf"):
            return jsonify({"error": "Only PDF files accepted"}), 400

        tmp_path = None
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            file.save(tmp.name)
            tmp_path = tmp.name

        try:
            from pdf2image import convert_from_path
            images = convert_from_path(tmp_path, dpi=150, first_page=1, last_page=3)

            image_contents = []
            for img in images:
                buf = BytesIO()
                img.save(buf, format="JPEG", quality=85)
                b64 = base64.b64encode(buf.getvalue()).decode()
                image_contents.append({
                    "type": "image_url",
                    "image_url": {"url": f"data:image/jpeg;base64,{b64}", "detail": "high"}
                })

            client = openai.OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
            prompt = """Extract contract data from these PDF pages. Return JSON only:
{
  "company_name": "...",
  "items": [
    {"product_name": "...", "net_price": 0, "cost": 0, "notes": "..."}
  ]
}"""
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[{
                    "role": "user",
                    "content": [{"type": "text", "text": prompt}] + image_contents
                }],
                max_tokens=2000
            )

            raw = response.choices[0].message.content.strip()
            json_match = re.search(r'\{[\s\S]*\}', raw)
            if json_match:
                result = json.loads(json_match.group())
            else:
                result = json.loads(raw)

            return jsonify(result)

        finally:
            if tmp_path:
                os.unlink(tmp_path)

    except Exception as e:
        import traceback
        return jsonify({"error": repr(e), "type": type(e).__name__, "trace": traceback.format_exc()}), 500


@app.route("/api/import-sheets", methods=["POST"])
def import_sheets():
    """Import extracted contract data to Google Sheets."""
    import traceback
    try:
        data = request.json or {}
        items = data.get("items", [])
        company = data.get("company_name", "")
        spreadsheet_id = data.get("spreadsheet_id", SPREADSHEET_ID)

        if not items:
            return jsonify({"error": "No items provided"}), 400

        creds_json_str = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
        if not creds_json_str:
            return jsonify({"error": "GOOGLE_CREDENTIALS_JSON not set"}), 500

        import gspread
        from google.oauth2.service_account import Credentials

        creds_info = json.loads(creds_json_str)
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(spreadsheet_id)

        ws = None
        for worksheet in sh.worksheets():
            if worksheet.id == SHEET_GID:
                ws = worksheet
                break
        if ws is None:
            ws = sh.sheet1

        rows = []
        for item in items:
            rows.append([
                company,
                item.get("product_name", ""),
                item.get("net_price", 0),
                item.get("cost", 0),
                item.get("notes", "")
            ])

        ws.append_rows(rows)
        return jsonify({"success": True, "rows_added": len(rows)})

    except Exception as e:
        return jsonify({"error": repr(e), "type": type(e).__name__, "trace": traceback.format_exc()}), 500


@app.route("/api/debug-sa", methods=["GET"])
def debug_sa():
    """Return service account email for debugging."""
    try:
        creds_raw = os.environ.get("GOOGLE_CREDENTIALS_JSON", "{}")
        creds = json.loads(creds_raw)
        return jsonify({
            "client_email": creds.get("client_email", "not found"),
            "project_id": creds.get("project_id", "not found")
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/debug-sheets", methods=["GET"])
def debug_sheets():
    """Step-by-step Google Sheets connection diagnostics."""
    import gspread, traceback
    from google.oauth2.service_account import Credentials
    steps = []
    try:
        creds_info = json.loads(os.environ.get("GOOGLE_CREDENTIALS_JSON", "{}"))
        steps.append("parsed credentials ok")
        creds = Credentials.from_service_account_info(creds_info, scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ])
        steps.append("created Credentials object")
        gc = gspread.authorize(creds)
        steps.append("gspread.authorize ok")
        sh = gc.open_by_key(SPREADSHEET_ID)
        steps.append("opened spreadsheet: " + sh.title)
        ws_list = sh.worksheets()
        steps.append("worksheets: " + str([(w.title, w.id) for w in ws_list]))
        return jsonify({"ok": True, "steps": steps})
    except Exception as e:
        return jsonify({"ok": False, "steps": steps, "error": repr(e), "trace": traceback.format_exc()})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port, debug=False)
