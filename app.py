#!/usr/bin/env python3
"""
Contract Data Importer â Backend (Railway)
Love Andaman
"""

from flask import Flask, request, jsonify
from flask_cors import CORS
import os
import json
import re
import tempfile
import base64
from io import BytesIO

app = Flask(__name__)

# CORS â à¸­à¸à¸¸à¸à¸²à¸ frontend Vercel à¹à¸£à¸µà¸¢à¸ API à¹à¸à¹
CORS(app, origins="*")

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1KWqJVYfoaRg3DwslW2zSQmPgScPbE9Z-0v-Ijwtdpms")
SHEET_GID = int(os.environ.get("SHEET_GID", "384942453"))


# âââ Health Check âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

@app.route("/")
def health():
    return jsonify({
        "status": "ok",
        "service": "Contract Importer API â Love Andaman",
        "has_api_key": bool(os.environ.get("ANTHROPIC_API_KEY")),
        "has_credentials": bool(os.environ.get("GOOGLE_CREDENTIALS_JSON"))
    })


@app.route("/api/status")
def status():
    return jsonify({
        "has_api_key": bool(os.environ.get("ANTHROPIC_API_KEY")),
        "has_credentials": bool(os.environ.get("GOOGLE_CREDENTIALS_JSON")),
        "spreadsheet_id": SPREADSHEET_ID
    })


# âââ Extract ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

@app.route("/api/extract", methods=["POST"])
def extract():
    if "pdf" not in request.files:
        return jsonify({"error": "à¹à¸¡à¹à¸à¸à¹à¸à¸¥à¹ PDF"}), 400

    pdf_file = request.files["pdf"]
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")

    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        pdf_file.save(tmp.name)
        tmp_path = tmp.name

    try:
        from pdf2image import convert_from_path
        images = convert_from_path(tmp_path, dpi=150)

        if api_key:
            result = extract_with_claude(images, api_key)
        else:
            result = extract_with_ocr(images)

        return jsonify(result)

    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        os.unlink(tmp_path)


def extract_with_claude(images, api_key):
    import anthropic

    client = anthropic.Anthropic(api_key=api_key)
    content = []

    for img in images[:4]:
        buffer = BytesIO()
        img.save(buffer, format="PNG")
        img_data = base64.b64encode(buffer.getvalue()).decode()
        content.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": "image/png",
                "data": img_data
            }
        })
    content.append({"type": "text", "text": """


à¹à¸à¸à¸µà¹à¹à¸¥à¸°à¸à¸­à¸à¸à¸¥à¸±à¸à¹à¸à¹à¸ JSON à¹à¸à¹à¸²à¸à¸±à¹à¸ à¹à¸¡à¹à¸¡à¸µà¸à¹à¸­à¸à¸§à¸²à¸¡à¸­à¸·à¹à¸:

{
  "company_name": "à¸à¸·à¹à¸­à¸à¸£à¸´à¸©à¸±à¸à¸à¸¹à¹à¹à¸«à¹à¸à¸£à¸´à¸à¸²à¸£à¸à¸±à¸§à¸£à¹ (operator/supplier à¹à¸¡à¹à¹à¸à¹ travel agent)",
  "items": [
    {
      "product_name": "à¸à¸·à¹à¸­à¹à¸à¸£à¹à¸à¸£à¸¡à¸à¸±à¸§à¸£à¹/à¸ªà¸´à¸à¸à¹à¸²",
      "net_price": 1000,
      "cost": 1000,
      "notes": "à¸«à¸¡à¸²à¸¢à¹à¸«à¸à¸¸ à¹à¸à¹à¸ Adult/Child, à¸à¸³à¸à¸§à¸à¸à¸, à¸«à¸¡à¸§à¸à¸«à¸¡à¸¹à¹"
    }
  ]
}

à¸à¸à¸à¸²à¸£à¸à¸¶à¸à¸à¹à¸­à¸¡à¸¹à¸¥:
- company_name: à¸à¸£à¸´à¸©à¸±à¸à¸à¸¹à¹à¹à¸«à¹à¸à¸£à¸´à¸à¸²à¸£à¸à¸±à¸§à¸£à¹ (à¹à¸¡à¹à¹à¸à¹ travel agent à¸«à¸£à¸·à¸­ agent à¸à¸µà¹à¸ªà¹à¸ contract à¸¡à¸²)
- product_name: à¸à¸·à¹à¸­à¸à¸±à¸§à¸£à¹/à¹à¸à¸£à¹à¸à¸£à¸¡à¹à¸à¹à¸¥à¸°à¸£à¸²à¸¢à¸à¸²à¸£à¸­à¸¢à¹à¸²à¸à¸à¸±à¸à¹à¸à¸
- net_price: à¸£à¸²à¸à¸² NET à¸à¸µà¹ agent à¸à¹à¸²à¸¢ (à¸à¸±à¸§à¹à¸¥à¸ THB à¸­à¸¢à¹à¸²à¸à¹à¸à¸µà¸¢à¸§ à¹à¸¡à¹à¸¡à¸µà¸«à¸à¹à¸§à¸¢)
- cost: à¹à¸«à¸¡à¸·à¸­à¸ net_price à¸«à¸²à¸à¹à¸¡à¹à¸¡à¸µ cost column à¹à¸¢à¸à¸à¹à¸²à¸à¸«à¸²à¸
- à¸£à¸§à¸¡à¸à¸¸à¸à¸ªà¸´à¸à¸à¹à¸²/à¹à¸à¸£à¹à¸à¸£à¸¡à¸à¸µà¹à¸¡à¸µà¹à¸à¹à¸­à¸à¸ªà¸²à¸£ (Adult, Child, Infant, à¸à¹à¸²à¸à¸à¸³à¸à¸§à¸à¸à¸)
- à¸«à¸²à¸à¸£à¸²à¸à¸²à¹à¸à¸à¸à¹à¸²à¸à¸à¸²à¸¡à¸à¸³à¸à¸§à¸à¸à¸ à¹à¸«à¹à¹à¸¢à¸à¹à¸à¹à¸ items à¸à¸£à¹à¸­à¸¡à¸£à¸°à¸à¸¸à¹à¸ notes
- à¸à¸­à¸à¸à¸¥à¸±à¸à¹à¸à¹à¸ JSON à¸­à¸¢à¹à¸²à¸à¹à¸à¸µà¸¢à¸§ à¹à¸¡à¹à¸¡à¸µ markdown code block"""
    })

    response = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=4000,
        messages=[{"role": "user", "content": content}]
    )

    text = response.content[0].text.strip()
    text = re.sub(r"^```(?:json)?\s*\n?", "", text, flags=re.MULTILINE)
    text = re.sub(r"\n?```\s*$", "", text, flags=re.MULTILINE).strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", text, re.DOTALL)
        if match:
            return json.loads(match.group())
        raise ValueError(f"à¹à¸¡à¹à¸ªà¸²à¸¡à¸²à¸£à¸à¹à¸à¸¥à¸ JSON: {text[:300]}")


def extract_with_ocr(images):
    try:
        import pytesseract
    except ImportError:
        return {
            "company_name": "",
            "items": [],
            "warning": "â ï¸ à¹à¸¡à¹à¸à¸ ANTHROPIC_API_KEY â à¸à¸£à¸¸à¸à¸²à¸à¸±à¹à¸à¸à¹à¸² Environment Variable à¸à¸ Railway"
        }

    full_text = ""
    for img in images:
        full_text += pytesseract.image_to_string(img, lang="eng") + "\n"

    company = ""
    for pattern in [r"Operator name[:\s]+([^\n\r]+)", r"à¸à¸£à¸´à¸©à¸±à¸[:\s]+([^\n\r]+)"]:
        m = re.search(pattern, full_text, re.IGNORECASE)
        if m:
            company = m.group(1).strip()
            break

    items = []
    price_re = re.compile(r"(\d{1,3}(?:,\d{3})+|\d{4,6})")
    for line in full_text.split("\n"):
        line = line.strip()
        prices = price_re.findall(line)
        if prices:
            product = price_re.sub("", line).strip(" .,:-")
            product = re.sub(r"\s+", " ", product)
            if product and len(product) > 2:
                net = int(prices[0].replace(",", ""))
                items.append({"product_name": product, "net_price": net, "cost": net, "notes": ""})

    return {
        "company_name": company,
        "items": items[:30],
        "warning": "â ï¸ à¹à¸à¹ OCR à¸à¸£à¸£à¸¡à¸à¸² à¸à¸£à¸¸à¸à¸²à¸à¸£à¸§à¸à¸ªà¸­à¸à¸à¹à¸­à¸¡à¸¹à¸¥à¸à¹à¸­à¸à¸à¸³à¹à¸à¹à¸²"
    }


# âââ Import to Sheets âââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

@app.route("/api/import-sheets", methods=["POST"])
def import_sheets():
    data = request.json or {}
    items = data.get("items", [])
    company = data.get("company_name", "")
    spreadsheet_id = data.get("spreadsheet_id", SPREADSHEET_ID)

    if not items:
        return jsonify({"error": "à¹à¸¡à¹à¸¡à¸µà¸à¹à¸­à¸¡à¸¹à¸¥à¸à¸µà¹à¸à¸°à¸à¸³à¹à¸à¹à¸²"}), 400

    # à¸­à¹à¸²à¸ credentials à¸à¸²à¸ Environment Variable
    creds_json_str = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
    if not creds_json_str:
        return jsonify({
            "error": "à¹à¸¡à¹à¸à¸ GOOGLE_CREDENTIALS_JSON",
            "help": "à¸à¸£à¸¸à¸à¸²à¸à¸±à¹à¸à¸à¹à¸² Environment Variable GOOGLE_CREDENTIALS_JSON à¸à¸ Railway"
        }), 400

    try:
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

        # à¸«à¸² worksheet à¸à¸²à¸¡ GID
        ws = None
        for worksheet in sh.worksheets():
            if worksheet.id == SHEET_GID:
                ws = worksheet
                break
        if ws is None:
            ws = sh.sheet1

        # Append rows
        rows = []
        for item in items:
            rows.append([
                company,
                item.get("product_name", ""),
                item.get("net_price", ""),
                item.get("cost", ""),
                item.get("notes", "")
            ])

        ws.append_rows(rows, value_input_option="USER_ENTERED")

        return jsonify({
            "success": True,
            "rows_added": len(rows),
            "message": f"à¸à¸³à¹à¸à¹à¸²à¸à¹à¸­à¸¡à¸¹à¸¥à¸ªà¸³à¹à¸£à¹à¸ {len(rows)} à¸£à¸²à¸¢à¸à¸²à¸£"
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# âââ Start ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
