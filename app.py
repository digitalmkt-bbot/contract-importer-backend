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
        "has_api_key": bool(os.environ.get("OPENAI_API_KEY")),
        "has_credentials": bool(os.environ.get("GOOGLE_CREDENTIALS_JSON"))
    })


@app.route("/api/status")
def status():
    creds_str = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
    service_account_email = ""
    if creds_str:
        try:
            service_account_email = json.loads(creds_str).get("client_email", "")
        except Exception:
            pass
    return jsonify({
        "has_api_key": bool(os.environ.get("OPENAI_API_KEY")),
        "has_credentials": bool(creds_str),
        "spreadsheet_id": SPREADSHEET_ID,
        "service_account_email": service_account_email
    })


# ─── Extract ──────────────────────────────────────────────────────────────────

@app.route("/api/extract", methods=["POST"])
def extract():
    # รับทั้ง "file" (ใหม่) และ "pdf" (เก่า) เพื่อ backward-compatibility
    uploaded = request.files.get("file") or request.files.get("pdf")
    if not uploaded:
        return jsonify({"error": "ไม่พบไฟล์ (ส่งเป็น field ชื่อ 'file')"}), 400

    api_key = os.environ.get("OPENAI_API_KEY", "")
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

        print(f"[EXTRACT] file={filename!r} is_pdf={is_pdf} is_image={is_image} has_key={bool(api_key)}", flush=True)

        if is_pdf:
            from pdf2image import convert_from_path
            images = convert_from_path(tmp_path, dpi=150)
            print(f"[EXTRACT] PDF pages={len(images)}", flush=True)
        elif is_image:
            img = PILImage.open(tmp_path).convert("RGB")
            images = [img]
            print(f"[EXTRACT] Image size={img.size}", flush=True)
        else:
            return jsonify({"error": "รองรับเฉพาะไฟล์ PDF, PNG, JPG, JPEG, WEBP เท่านั้น"}), 400

        if api_key:
            result = extract_with_claude(images, api_key)
        else:
            result = extract_with_ocr(images)

        print(f"[EXTRACT] done: company={result.get('company_name')!r} items={len(result.get('items',[]))}", flush=True)
        return jsonify(result)

    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500
    finally:
        os.unlink(tmp_path)


def extract_with_claude(images, api_key):
    from openai import OpenAI

    client = OpenAI(api_key=api_key)

    system_prompt = (
        "You are a data extraction assistant for a travel agency's internal pricing system. "
        "Your job is to read tour operator rate sheets / price contracts and extract structured pricing data. "
        "Always respond with valid JSON only — no markdown, no explanations."
    )

    user_prompt = """Extract ALL pricing data from this tour operator contract/rate sheet image(s).
CRITICAL: You MUST extract EVERY SINGLE ROW that contains a price or rate — do NOT skip, summarize, or truncate any row.

Return ONLY a JSON object in this exact format (no markdown code blocks, no extra text):
{
  "company_name": "name of the tour operator / supplier company",
  "items": [
    {
      "product_name": "Oneday Tour Type | Program Name | From → To | DEP. Time - ARR. Time | Passenger Type",
      "net_rate": 1000,
      "selling_rate": 1500,
      "notes": "any other remarks or conditions"
    }
  ]
}

Rules:
- company_name: the operator/supplier name (not the travel agent)
- product_name: combine ALL available fields below, separated by " | " — include every piece of information found:
    1. Tour type / category (e.g. "Oneday Tour", "Speedboat Tour", "Liveaboard", "Transfer", "Package") — include if shown
    2. Program name (e.g. "Phi Phi Island Tour", "Similan Diving Day Trip", "ATV Adventure") — always include
    3. Route: departure point → destination (e.g. "Phuket → Phi Phi", "Khao Lak → Similan", "Krabi → Railay") — look for columns: From/To, From, Route, Departure Point, Origin/Destination — include if found
    4. Departure & arrival time (e.g. "DEP. 07:30 - ARR. 17:00", "08:00-17:00", "Full Day", "Half Day AM") — look for columns: Time DEP.-ARR., Time, Schedule, Departure Time — include if found
    5. PASSENGER TYPE / AGE RANGE — ALWAYS append as the LAST segment of product_name whenever the row has a specific passenger type or age category:
       - Use the exact label from the document: "Adult", "Child", "Infant", "CHD", "INF", "Pax 1-4", "Min 15 Pax", "Per Person", etc.
       - This is MANDATORY: every row that corresponds to a specific age group or pax tier MUST have that label at the end of product_name.
       - Do NOT put passenger type / age category in the notes field.
    Build product_name by joining only the fields that are actually present in the document.
    Examples:
      "Oneday Tour | Phi Phi Island Tour | Phuket → Phi Phi | DEP. 07:30 - ARR. 17:00 | Adult"
      "Oneday Tour | Phi Phi Island Tour | Phuket → Phi Phi | DEP. 07:30 - ARR. 17:00 | Child"
      "Oneday Tour | Phi Phi Island Tour | Phuket → Phi Phi | DEP. 07:30 - ARR. 17:00 | Infant"
      "Speedboat Tour | Similan Day Trip | Khao Lak → Similan | DEP. 06:00 - ARR. 18:00 | Adult"
      "Liveaboard | Similan Islands | Khao Lak → Similan | 2D1N | Adult"
      "Transfer | Phuket Airport → Patong Hotel | 1-4 Pax"
      "ATV Adventure | Chalong Circuit | 09:00-12:00 | Adult"
    Include as much detail as possible — do NOT omit route, time, or passenger type if they appear anywhere in the row or header.
- net_rate: agent/net cost price in THB (number only). Look for: Net Rate, Net Price, Agent Rate, Net, Cost
- selling_rate: retail/public price in THB. Look for: Selling Rate, Public Rate, Rack Rate, Adult Rate, Full Price. Use 0 if not found.
- notes: CRITICAL — copy the COMPLETE, FULL, VERBATIM text from all remark/note/condition fields for each row. Do NOT summarize, shorten, or omit any part of the remark text. Extract and combine ALL of the following:
    1. COPY THE ENTIRE TEXT from any column named "Remark", "Remarks", "หมายเหตุ", "Note", "Notes", "Condition", "Conditions", "Remark/Note" — paste every word exactly as written, including numbers, symbols, and line breaks (use space to join multiple lines)
    2. "Extra Transfer" / "Extra Transfer Fee" — if the document has a section, row, or column for extra transfer cost/conditions relating to a product, copy that full value into the notes of the matching product row (match by product name, tour type, or proximity in the table)
    3. Any other special conditions NOT related to passenger type/age group (e.g. "Include transfer", "Min 2 pax", "Seasonal surcharge applies", "Valid Nov-Apr")
    If multiple remark values exist for a row, join them with " | ". If the remark cell is empty or the column does not exist, leave notes as empty string.
    IMPORTANT: Every row that has ANY text in a Remark/Note/Condition column MUST have that text in its notes field — never leave notes empty when remark data exists.
- MUST include ALL line items without exception: Adult, Child, Infant, every pax count variation, every category
- If prices vary by passenger type or group size, each must be a SEPARATE item — with the type/tier appended to product_name
- If a table header applies to multiple rows below it, repeat the header info in each row's product_name
- Do NOT stop early — extract until the LAST row of data on the page
- Return ONLY the JSON object"""

    # Process 1 page per batch — maximum token budget per page for complete extraction
    all_items = []
    company_name = ""
    page_batches = [images[i:i+1] for i in range(0, len(images), 1)]

    for batch_idx, batch in enumerate(page_batches):
        content = []
        for img in batch:
            buffer = BytesIO()
            img.save(buffer, format="JPEG", quality=85)
            img_data = base64.b64encode(buffer.getvalue()).decode()
            content.append({
                "type": "image_url",
                "image_url": {
                    "url": f"data:image/jpeg;base64,{img_data}",
                    "detail": "high"
                }
            })
        content.append({"type": "text", "text": user_prompt})

        response = client.chat.completions.create(
            model="gpt-4o-2024-08-06",
            max_tokens=16000,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": content}
            ]
        )

        text = response.choices[0].message.content.strip()
        finish_reason = response.choices[0].finish_reason
        print(f"[GPT] batch={batch_idx} finish={finish_reason!r} len={len(text)} preview={text[:100]!r}", flush=True)

        # Detect refusal
        refusal_phrases = ["i'm sorry", "i cannot", "i can't", "unable to assist", "can't assist", "cannot assist"]
        if any(p in text.lower() for p in refusal_phrases) and "{" not in text:
            print(f"[GPT] REFUSAL detected batch {batch_idx}", flush=True)
            continue

        # Strip markdown code fences
        text = re.sub(r"^```(?:json)?\s*\n?", "", text, flags=re.MULTILINE)
        text = re.sub(r"\n?```\s*$", "", text, flags=re.MULTILINE).strip()

        try:
            batch_result = json.loads(text)
            print(f"[GPT] JSON OK: {len(batch_result.get('items',[]))} items", flush=True)
        except json.JSONDecodeError as je:
            print(f"[GPT] JSON FAIL: {je} tail={text[-80:]!r}", flush=True)
            match = re.search(r"\{.*\}", text, re.DOTALL)
            if match:
                try:
                    batch_result = json.loads(match.group())
                    print(f"[GPT] regex fallback OK: {len(batch_result.get('items',[]))} items", flush=True)
                except Exception as e2:
                    print(f"[GPT] regex fallback FAIL: {e2}", flush=True)
                    continue
            else:
                print(f"[GPT] no JSON found, skipping batch {batch_idx}", flush=True)
                continue

        if batch_idx == 0 and not company_name:
            company_name = batch_result.get("company_name", "")
        all_items.extend(batch_result.get("items", []))

    # Deduplicate items by product_name + net_rate
    seen = set()
    deduped = []
    for item in all_items:
        key = (item.get("product_name", ""), str(item.get("net_rate", "")))
        if key not in seen:
            seen.add(key)
            deduped.append(item)

    return {"company_name": company_name, "items": deduped}


def _extract_single_batch(client, system_prompt, user_prompt, images):
    """Legacy single-batch extraction (kept for reference)."""
    content = []
    for img in images:
        buffer = BytesIO()
        img.save(buffer, format="JPEG", quality=85)
        img_data = base64.b64encode(buffer.getvalue()).decode()
        content.append({
            "type": "image_url",
            "image_url": {
                "url": f"data:image/jpeg;base64,{img_data}",
                "detail": "high"
            }
        })
    content.append({"type": "text", "text": user_prompt})

    response = client.chat.completions.create(
        model="gpt-4o",
        max_tokens=16000,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": content}
        ]
    )

    text = response.choices[0].message.content.strip()

    # Detect refusal
    refusal_phrases = ["i'm sorry", "i cannot", "i can't", "unable to assist", "can't assist", "cannot assist"]
    if any(p in text.lower() for p in refusal_phrases) and "{" not in text:
        raise ValueError(f"AI ไม่สามารถอ่านเอกสารนี้ได้ กรุณาลองใหม่หรือใช้ไฟล์ PNG/JPG แทน PDF")

    # Strip markdown code fences
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
            "warning": "⚠️ ไม่พบ OPENAI_API_KEY — กรุณาตั้งค่า Environment Variable บน Railway"
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

        # Sheet column layout: E=Operator, F=List/Tour name, G=empty, H=Net Rate, I=Selling Rate, J=Profit(formula),
        # K=Profit(%), L=Margin%, M=10%Commission, N-P=other cols, Q=Notes (หมายเหตุ)
        # NOTE: use ws.update() with explicit range E:Q — NOT append_rows()
        #       because append_rows() detects the table starting at col E and offsets all data by 4 cols.

        # Read all existing data for dedup + finding next empty row
        existing_keys = set()
        all_values = ws.get_all_values()
        for row in all_values[4:]:  # skip header rows 1-4
            e = row[4].strip() if len(row) > 4 else ""   # col E = Operator / company
            f = row[5].strip() if len(row) > 5 else ""   # col F = Product Name
            k = (e + "|" + f).lower()
            if k != "|":
                existing_keys.add(k)

        # Find the next empty row (scan from bottom, look for last row with E or F data)
        next_row = 5  # default: start at row 5 (after 4 header rows)
        for i in range(len(all_values) - 1, 3, -1):
            e = all_values[i][4].strip() if len(all_values[i]) > 4 else ""
            f = all_values[i][5].strip() if len(all_values[i]) > 5 else ""
            if e or f:
                next_row = i + 2  # i is 0-indexed; +1 for 1-indexed sheet, +1 for next row
                break

        # Filter: only add items not already in sheet
        rows = []
        skipped = []
        for item in items:
            name = item.get("product_name", "")
            k = (company.strip() + "|" + name.strip()).lower()
            if k in existing_keys:
                skipped.append(name)
            else:
                # Write directly to E:Q — NO A-D padding needed when using ws.update()
                rows.append([
                    company,   # col E = Operator name
                    name,      # col F = Product / tour name
                    "",        # col G (empty)
                    item.get("net_rate", item.get("net_price", "")),   # col H = Net Rate
                    item.get("selling_rate", item.get("public_rate", item.get("cost", ""))),  # col I = Selling Rate
                    "",        # col J = Profit amount (leave blank — formula fills this)
                    "",        # col K = Profit% (leave blank — formula fills this)
                    "",        # col L = Margin% (leave blank — formula fills this)
                    "",        # col M = 10% Commission (leave blank — formula fills this)
                    "",        # col N (empty)
                    "",        # col O (empty)
                    "",        # col P (empty)
                    item.get("notes", "")  # col Q = Notes (หมายเหตุ)
                ])
                existing_keys.add(k)

        if rows:
            end_row = next_row + len(rows) - 1
            ws.update(f"E{next_row}:Q{end_row}", rows, value_input_option="USER_ENTERED")

        skip_msg = f" (ข้ามซ้ำ {len(skipped)} รายการ)" if skipped else ""
        return jsonify({
            "success": True,
            "rows_added": len(rows),
            "rows_skipped": len(skipped),
            "message": f"✅ นำเข้าข้อมูลสำเร็จ {len(rows)} รายการ{skip_msg}"
        })

    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500


# ─── Start ────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
