#!/usr/bin/env python3
"""
Contract Data Importer â Backend (Railway)
Love Andaman
"""

from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import os
import json
import re
import tempfile
import base64
from io import BytesIO

app = Flask(__name__, static_folder="static", static_url_path="")

# CORS â à¸­à¸à¸¸à¸à¸²à¸ frontend Vercel à¹à¸£à¸µà¸¢à¸ API à¹à¸à¹
CORS(app, origins="*")

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1KWqJVYfoaRg3DwslW2zSQmPgScPbE9Z-0v-Ijwtdpms")
SHEET_GID = int(os.environ.get("SHEET_GID", "384942453"))


# âââ Health Check âââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

@app.route("/")
def index():
    # Serve frontend UI if it exists, otherwise show API status
    static_index = os.path.join(app.static_folder or "", "index.html")
    if os.path.exists(static_index):
        return send_from_directory(app.static_folder, "index.html")
    return jsonify({
        "status": "ok",
        "service": "Contract Importer API â Love Andaman",
        "has_api_key": bool(os.environ.get("OPENAI_API_KEY")),
        "has_credentials": bool(os.environ.get("GOOGLE_CREDENTIALS_JSON"))
    })


@app.route("/api/status")
def status():
    return jsonify({
        "has_api_key": bool(os.environ.get("OPENAI_API_KEY")),
        "has_credentials": bool(os.environ.get("GOOGLE_CREDENTIALS_JSON")),
        "spreadsheet_id": SPREADSHEET_ID
    })


# âââ Extract ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

@app.route("/api/extract", methods=["POST"])
def extract():
    # à¸£à¸±à¸à¸à¸±à¹à¸ "file" (à¹à¸«à¸¡à¹) à¹à¸¥à¸° "pdf" (à¹à¸à¹à¸²) à¹à¸à¸·à¹à¸­ backward-compatibility
    uploaded = request.files.get("file") or request.files.get("pdf")
    if not uploaded:
        return jsonify({"error": "à¹à¸¡à¹à¸à¸à¹à¸à¸¥à¹ (à¸ªà¹à¸à¹à¸à¹à¸ field à¸à¸·à¹à¸­ 'file')"}), 400

    api_key = os.environ.get("OPENAI_API_KEY", "")
    filename = (uploaded.filename or "").lower()
    is_pdf = filename.endswith(".pdf") or uploaded.content_type == "application/pdf"
    is_image = any(filename.endswith(ext) for ext in (".png", ".jpg", ".jpeg", ".webp"))

    # à¸à¸±à¸à¸à¸¶à¸ temp file
    suffix = ".pdf" if is_pdf else os.path.splitext(filename)[1] or ".png"
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        uploaded.save(tmp.name)
        tmp_path = tmp.name

    try:
        from PIL import Image as PILImage

        if is_pdf:
            from pdf2image import convert_from_path
            images = convert_from_path(tmp_path, dpi=150)
        elif is_image:
            img = PILImage.open(tmp_path).convert("RGB")
            images = [img]
        else:
            return jsonify({"error": "à¸£à¸­à¸à¸£à¸±à¸à¹à¸à¸à¸²à¸°à¹à¸à¸¥à¹ PDF, PNG, JPG, JPEG, WEBP à¹à¸à¹à¸²à¸à¸±à¹à¸"}), 400

        if api_key:
            result = extract_with_openai(images, api_key)
        else:
            result = extract_with_ocr(images)

        return jsonify(result)

    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500
    finally:
        os.unlink(tmp_path)


def extract_with_openai(images, api_key):
    from openai import OpenAI

    client = OpenAI(api_key=api_key)
    content = []

    prompt = """à¸à¸µà¹à¸à¸·à¸­à¹à¸­à¸à¸ªà¸²à¸£à¸ªà¸±à¸à¸à¸²à¸£à¸²à¸à¸²à¸à¸£à¸´à¸©à¸±à¸à¸à¸³à¹à¸à¸µà¹à¸¢à¸§ à¹à¸à¸£à¸à¸à¸¶à¸à¸à¹à¸­à¸¡à¸¹à¸¥à¸à¹à¸­à¹à¸à¸à¸µà¹à¹à¸¥à¸°à¸à¸­à¸à¸à¸¥à¸±à¸à¹à¸à¹à¸ JSON à¹à¸à¹à¸²à¸à¸±à¹à¸ à¹à¸¡à¹à¸¡à¸µà¸à¹à¸­à¸à¸§à¸²à¸¡à¸­à¸·à¹à¸:

{
  "company_name": "à¸à¸·à¹à¸­à¸à¸£à¸´à¸©à¸±à¸à¸à¸¹à¹à¹à¸«à¹à¸à¸£à¸´à¸à¸²à¸£à¸à¸±à¸§à¸£à¹ (operator/supplier à¹à¸¡à¹à¹à¸à¹ travel agent)",
  "items": [
    {
      "product_name": "à¸à¸·à¹à¸­à¹à¸à¸£à¹à¸à¸£à¸¡à¸à¸±à¸§à¸£à¹/à¸ªà¸´à¸à¸à¹à¸²",
      "net_rate": 1000,
      "selling_rate": 1500,
      "notes": "à¸«à¸¡à¸²à¸¢à¹à¸«à¸à¸¸ à¹à¸à¹à¸ Adult/Child, à¸à¸³à¸à¸§à¸à¸à¸, à¸«à¸¡à¸§à¸à¸«à¸¡à¸¹à¹"
    }
  ]
}

à¸à¸à¸à¸²à¸£à¸à¸¶à¸à¸à¹à¸­à¸¡à¸¹à¸¥:
- company_name: à¸à¸£à¸´à¸©à¸±à¸à¸à¸¹à¹à¹à¸«à¹à¸à¸£à¸´à¸à¸²à¸£à¸à¸±à¸§à¸£à¹ (à¹à¸¡à¹à¹à¸à¹ travel agent à¸«à¸£à¸·à¸­ agent à¸à¸µà¹à¸ªà¹à¸ contract à¸¡à¸²)
- product_name: à¸à¸·à¹à¸­à¸à¸±à¸§à¸£à¹/à¹à¸à¸£à¹à¸à¸£à¸¡à¹à¸à¹à¸¥à¸°à¸£à¸²à¸¢à¸à¸²à¸£à¸­à¸¢à¹à¸²à¸à¸à¸±à¸à¹à¸à¸
- net_rate: à¸£à¸²à¸à¸² NET à¸à¸µà¹ agent à¸à¹à¸²à¸¢ (à¸à¸±à¸§à¹à¸¥à¸ THB à¸­à¸¢à¹à¸²à¸à¹à¸à¸µà¸¢à¸§ à¹à¸¡à¹à¸¡à¸µà¸«à¸à¹à¸§à¸¢) â à¸­à¸²à¸à¹à¸à¹à¸à¸·à¹à¸­à¹à¸à¹à¸­à¸à¸ªà¸²à¸£à¸§à¹à¸² "Net Rate", "Net Price", "Agent Rate", "Cost"
- selling_rate: à¸£à¸²à¸à¸²à¸à¸²à¸¢à¹à¸«à¹à¸¥à¸¹à¸à¸à¹à¸² (à¸à¸±à¸§à¹à¸¥à¸ THB) â à¸­à¸²à¸à¹à¸à¹à¸à¸·à¹à¸­à¹à¸à¹à¸­à¸à¸ªà¸²à¸£à¸§à¹à¸² "Selling Rate", "Cost Rate", "Public Rate", "Rack Rate", "Price", "Adult Rate" à¸«à¸²à¸à¹à¸¡à¹à¸¡à¸µà¹à¸à¹à¸­à¸à¸ªà¸²à¸£à¹à¸«à¹à¹à¸ªà¹ 0
- à¸£à¸§à¸¡à¸à¸¸à¸à¸ªà¸´à¸à¸à¹à¸²/à¹à¸à¸£à¹à¸à¸£à¸¡à¸à¸µà¹à¸¡à¸µà¹à¸à¹à¸­à¸à¸ªà¸²à¸£ (Adult, Child, Infant, à¸à¹à¸²à¸à¸à¸³à¸à¸§à¸à¸à¸)
- à¸«à¸²à¸à¸£à¸²à¸à¸²à¹à¸à¸à¸à¹à¸²à¸à¸à¸²à¸¡à¸à¸³à¸à¸§à¸à¸à¸ à¹à¸«à¹à¹à¸¢à¸à¹à¸à¹à¸ items à¸à¸£à¹à¸­à¸¡à¸£à¸°à¸à¸¸à¹à¸ notes
- à¸à¸­à¸à¸à¸¥à¸±à¸à¹à¸à¹à¸ JSON à¸­à¸¢à¹à¸²à¸à¹à¸à¸µà¸¢à¸§ à¹à¸¡à¹à¸¡à¸µ markdown code block"""

    for img in images[:4]:
        buffer = BytesIO()
        img.save(buffer, format="PNG")
        img_data = base64.b64encode(buffer.getvalue()).decode()
        content.append({
            "type": "image_url",
            "image_url": {
                "url": f"data:image/png;base64,{img_data}",
                "detail": "high"
            }
        })

    content.append({"type": "text", "text": prompt})

    response = client.chat.completions.create(
        model="gpt-4o",
        max_tokens=4000,
        messages=[{"role": "user", "content": content}]
    )

    text = response.choices[0].message.content.strip()
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
            "warning": "â ï¸ à¹à¸¡à¹à¸à¸ OPENAI_API_KEY â à¸à¸£à¸¸à¸à¸²à¸à¸±à¹à¸à¸à¹à¸² Environment Variable à¸à¸ Railway"
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

        ws = None
        for worksheet in sh.worksheets():
            if worksheet.id == SHEET_GID:
                ws = worksheet
                break
        if ws is None:
            ws = sh.sheet1

        existing_keys = set()
        all_values = ws.get_all_values()
        for row in all_values[4:]:
            e = row[4].strip() if len(row) > 4 else ""
            f = row[5].strip() if len(row) > 5 else ""
            k = (e + "|" + f).lower()
            if k != "|":
                existing_keys.add(k)

        rows = []
        skipped = []
        for item in items:
            name = item.get("product_name", "")
            k = (company.strip() + "|" + name.strip()).lower()
            if k in existing_keys:
                skipped.append(name)
            else:
                rows.append([
                    company,
                    name,
                    item.get("net_rate", item.get("net_price", "")),
                    item.get("selling_rate", item.get("public_rate", item.get("cost", ""))),
                    item.get("notes", "")
                ])
                existing_keys.add(k)

        if rows:
            ws.append_rows(rows, value_input_option="USER_ENTERED")

        skip_msg = f", à¸à¹à¸²à¸¡à¸à¹à¸³ {len(skipped)} à¸£à¸²à¸¢à¸à¸²à¸£" if skipped else ""
        return jsonify({
            "success": True,
            "rows_added": len(rows),
            "rows_skipped": len(skipped),
            "message": f"à¸à¸³à¹à¸à¹à¸²à¸à¹à¸­à¸¡à¸¹à¸¥à¸ªà¸³à¹à¸£à¹à¸ {len(rows)} à¸£à¸²à¸¢à¸à¸²à¸£{skip_msg}"
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# âââ Start ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
