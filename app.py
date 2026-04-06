#!/usr/bin/env python3
"""
Contract Data Importer — Backend (Railway)
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

# CORS — อนุญาต frontend Vercel เรียก API ได้
CORS(app, origins="*")

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1KWqJVYfoaRg3DwslW2zSQmPgScPbE9Z-0v-Ijwtdpms")
SHEET_GID = int(os.environ.get("SHEET_GID", "384942453"))


# ─── Health Check ─────────────────────────────────────────────────────────────

@app.route("/")
def index():
    # Serve frontend UI if it exists, otherwise show API status
    static_index = os.path.join(app.static_folder or "", "index.html")
    if os.path.exists(static_index):
        return send_from_directory(app.static_folder, "index.html")
    return jsonify({
        "status": "ok",
        "service": "Contract Importer API — Love Andaman",
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


# ─── Extract ──────────────────────────────────────────────────────────────────

@app.route("/api/extract", methods=["POST"])
def extract():
    # รับทั้ง "file" (ใหม่) และ "pdf" (เก่า) เพื่อ backward-compatibility
    uploaded = request.files.get("file") or request.files.get("pdf")
    if not uploaded:
        return jsonify({"error": "ไม่พบไฟล์ (ส่งเป็น field ชื่อ 'file')"}), 400

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    filename = (uploaded.filename or "").lower()
    is_pdf = filename.endswith(".pdf") or uploaded.content_type == "application/pdf"
    is_image = any(filename.endswith(ext) for ext in (".png", ".jpg", ".jpeg", ".webp"))

    # บันทึก temp file
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
            return jsonify({"error": "รองรับเฉพาะไฟล์ PDF, PNG, JPG, JPEG, WEBP เท่านั้น"}), 400

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

    content.append({
        "type": "text",
        "text": """นี่คือ PDF สัญญาราคาบริษัทนำเที่ยว โปรดดึงข้อมูลต่อไปนี้และตอบกลับเป็น JSON เท่านั้น ไม่มีข้อความอื่น:

{
  "company_name": "ชื่อบริษัทผู้ให้บริการทัวร์ (operator/supplier ไม่ใช่ travel agent)",
  "items": [
    {
      "product_name": "ชื่อโปรแกรมทัวร์/สินค้า",
      "net_rate": 1000,
      "selling_rate": 1500,
      "notes": "หมายเหตุ เช่น Adult/Child, จำนวนคน, หมวดหมู่"
    }
  ]
}

กฎการดึงข้อมูล:
- company_name: บริษัทผู้ให้บริการทัวร์ (ไม่ใช่ travel agent หรือ agent ที่ส่ง contract มา)
- product_name: ชื่อทัวร์/โปรแกรมแต่ละรายการอย่างชัดเจน
- net_rate: ราคา NET ที่ agent จ่าย (ตัวเลข THB อย่างเดียว ไม่มีหน่วย) — อาจใช้ชื่อในเอกสารว่า "Net Rate", "Net Price", "Agent Rate", "Cost"
- selling_rate: ราคาขายให้ลูกค้า (ตัวเลข THB) — อาจใช้ชื่อในเอกสารว่า "Selling Rate", "Cost Rate", "Public Rate", "Rack Rate", "Price", "Adult Rate" (ราคาที่มักอยู่คู่กับ Net Rate) หากไม่มีในเอกสารให้ใส่ 0
- รวมทุกสินค้า/โปรแกรมที่มีในเอกสาร (Adult, Child, Infant, ต่างจำนวนคน)
- หากราคาแตกต่างตามจำนวนคน ให้แยกเป็น items พร้อมระบุใน notes
- ตอบกลับเป็น JSON อย่างเดียว ไม่มี markdown code block"""
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
        raise ValueError(f"ไม่สามารถแปลง JSON: {text[:300]}")


def extract_with_ocr(images):
    try:
        import pytesseract
    except ImportError:
        return {
            "company_name": "",
            "items": [],
            "warning": "⚠️ ไม่พบ ANTHROPIC_API_KEY — กรุณาตั้งค่า Environment Variable บน Railway"
        }

    full_text = ""
    for img in images:
        full_text += pytesseract.image_to_string(img, lang="eng") + "\n"

    company = ""
    for pattern in [r"Operator name[:\s]+([^\n\r]+)", r"บริษัท[:\s]+([^\n\r]+)"]:
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
        "warning": "⚠️ ใช้ OCR ธรรมดา กรุณาตรวจสอบข้อมูลก่อนนำเข้า"
    }


# ─── Import to Sheets ─────────────────────────────────────────────────────────

@app.route("/api/import-sheets", methods=["POST"])
def import_sheets():
    data = request.json or {}
    items = data.get("items", [])
    company = data.get("company_name", "")
    spreadsheet_id = data.get("spreadsheet_id", SPREADSHEET_ID)

    if not items:
        return jsonify({"error": "ไม่มีข้อมูลที่จะนำเข้า"}), 400

    # อ่าน credentials จาก Environment Variable
    creds_json_str = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
    if not creds_json_str:
        return jsonify({
            "error": "ไม่พบ GOOGLE_CREDENTIALS_JSON",
            "help": "กรุณาตั้งค่า Environment Variable GOOGLE_CREDENTIALS_JSON บน Railway"
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

        # หา worksheet ตาม GID
        ws = None
        for worksheet in sh.worksheets():
            if worksheet.id == SHEET_GID:
                ws = worksheet
                break
        if ws is None:
            ws = sh.sheet1

        # Dedup: build set of existing (company|product_name) keys (rows 5+, cols E-F)
        existing_keys = set()
        all_values = ws.get_all_values()
        for row in all_values[4:]:  # skip header rows 1-4
            e = row[4].strip() if len(row) > 4 else ""
            f = row[5].strip() if len(row) > 5 else ""
            k = (e + "|" + f).lower()
            if k != "|":
                existing_keys.add(k)

        # Filter: only add items not already in sheet
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

        skip_msg = f", ข้ามซ้ำ {len(skipped)} รายการ" if skipped else ""
        return jsonify({
            "success": True,
            "rows_added": len(rows),
            "rows_skipped": len(skipped),
            "message": f"นำเข้าข้อมูลสำเร็จ {len(rows)} รายการ{skip_msg}"
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ─── Start ────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
