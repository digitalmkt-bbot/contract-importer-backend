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
import logging
from io import BytesIO

# ─── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=os.environ.get("LOG_LEVEL", "INFO").upper(),
    format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
)
log = logging.getLogger("contract-importer")

app = Flask(__name__, static_folder="static", static_url_path="")

# Upload limit (default 25 MB — enough for multi-page scanned PDFs)
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_UPLOAD_MB", "25")) * 1024 * 1024

# CORS — default allow-all for dev; restrict in prod via CORS_ORIGINS="https://foo.com,https://bar.com"
_cors_origins_env = os.environ.get("CORS_ORIGINS", "*").strip()
if _cors_origins_env and _cors_origins_env != "*":
    _cors_origins = [o.strip() for o in _cors_origins_env.split(",") if o.strip()]
else:
    _cors_origins = "*"
CORS(app, origins=_cors_origins)

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1X_gcLo3RROT11Hv9qvhiegoztk9STv4lP2aTq0Ih0Ho")
SHEET_GID = int(os.environ.get("SHEET_GID", "384942453"))
OPENAI_MODEL = os.environ.get("OPENAI_MODEL", "gpt-4o-2024-08-06")
OPENAI_MAX_TOKENS = int(os.environ.get("OPENAI_MAX_TOKENS", "16000"))
OPENAI_CONTINUATION_ROUNDS = int(os.environ.get("OPENAI_CONTINUATION_ROUNDS", "5"))

# Allowed spreadsheet ID pattern — Google Sheets IDs are 40-50 alphanumeric + - + _
_SHEET_ID_RE = re.compile(r"^[A-Za-z0-9_-]{20,120}$")

# ─── Auth + Rate Limit ────────────────────────────────────────────────────────
import time as _time
import threading as _threading
from collections import deque as _deque
from functools import wraps as _wraps

API_TOKEN = os.environ.get("API_TOKEN", "").strip()
RATE_LIMIT_MAX = int(os.environ.get("RATE_LIMIT_MAX", "10"))        # requests
RATE_LIMIT_WINDOW = int(os.environ.get("RATE_LIMIT_WINDOW", "60"))  # seconds

_rate_lock = _threading.Lock()
_rate_buckets: "dict[str, _deque]" = {}


def _client_ip(req):
    """Best-effort client IP that respects Railway/Vercel reverse proxies."""
    fwd = req.headers.get("X-Forwarded-For", "")
    if fwd:
        return fwd.split(",")[0].strip()
    return req.remote_addr or "unknown"


def _rate_limit_check(ip):
    """Return (ok, retry_after_seconds)."""
    if RATE_LIMIT_MAX <= 0:
        return True, 0
    now = _time.monotonic()
    window_start = now - RATE_LIMIT_WINDOW
    with _rate_lock:
        bucket = _rate_buckets.setdefault(ip, _deque())
        while bucket and bucket[0] < window_start:
            bucket.popleft()
        if len(bucket) >= RATE_LIMIT_MAX:
            retry = int(bucket[0] + RATE_LIMIT_WINDOW - now) + 1
            return False, max(retry, 1)
        bucket.append(now)
        return True, 0


def require_auth(func):
    """Decorator: enforce API_TOKEN (if set) + per-IP rate limit."""
    @_wraps(func)
    def wrapper(*args, **kwargs):
        if API_TOKEN:
            provided = (request.headers.get("X-API-Token")
                        or request.args.get("api_token")
                        or "")
            if provided != API_TOKEN:
                log.warning("Auth rejected for %s on %s", _client_ip(request), request.path)
                return jsonify({"error": "Unauthorized — missing or invalid X-API-Token"}), 401
        ok, retry = _rate_limit_check(_client_ip(request))
        if not ok:
            resp = jsonify({
                "error": f"Rate limit exceeded — try again in {retry}s",
                "retry_after": retry,
            })
            resp.status_code = 429
            resp.headers["Retry-After"] = str(retry)
            return resp
        return func(*args, **kwargs)
    return wrapper


@app.errorhandler(413)
def _too_large(_):
    limit_mb = app.config["MAX_CONTENT_LENGTH"] // (1024 * 1024)
    return jsonify({"error": f"ไฟล์ใหญ่เกินไป (จำกัด {limit_mb} MB)"}), 413


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
        "sheet_gid": SHEET_GID,
        "service_account_email": service_account_email,
        "model": OPENAI_MODEL,
        "max_upload_mb": app.config["MAX_CONTENT_LENGTH"] // (1024 * 1024),
        "cors_origins": _cors_origins if _cors_origins != "*" else "*",
        "auth_required": bool(API_TOKEN),
        "rate_limit": {"max": RATE_LIMIT_MAX, "window_seconds": RATE_LIMIT_WINDOW},
    })


# ─── Extract ──────────────────────────────────────────────────────────────────

@app.route("/api/extract", methods=["POST"])
@require_auth
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

        log.info(f"[EXTRACT] file={filename!r} content_type={uploaded.content_type!r} is_pdf={is_pdf} is_image={is_image} has_key={bool(api_key)}")
        if is_pdf:
            from pdf2image import convert_from_path
            images = convert_from_path(tmp_path, dpi=200)
            log.info(f"[EXTRACT] PDF pages={len(images)}")
        elif is_image:
            img = PILImage.open(tmp_path).convert("RGB")
            images = [img]
            log.info(f"[EXTRACT] Image size={img.size}")
        else:
            return jsonify({"error": "รองรับเฉพาะไฟล์ PDF, PNG, JPG, JPEG, WEBP เท่านั้น"}), 400

        if api_key:
            result = extract_with_claude(images, api_key)
        else:
            result = extract_with_ocr(images)

        log.info(f"[EXTRACT] done: company={result.get('company_name')!r} items={len(result.get('items',[]))}")
        return jsonify(result)

    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500
    finally:
        os.unlink(tmp_path)


@app.route("/api/extract/stream", methods=["POST"])
@require_auth
def extract_stream():
    """NDJSON streaming extract — one JSON object per line for real-time progress."""
    from flask import Response, stream_with_context

    uploaded = request.files.get("file") or request.files.get("pdf")
    if not uploaded:
        return jsonify({"error": "ไม่พบไฟล์ (ส่งเป็น field ชื่อ 'file')"}), 400

    api_key = os.environ.get("OPENAI_API_KEY", "")
    filename = (uploaded.filename or "").lower()
    is_pdf = filename.endswith(".pdf") or uploaded.content_type == "application/pdf"
    is_image = any(filename.endswith(ext) for ext in (".png", ".jpg", ".jpeg", ".webp"))

    suffix = ".pdf" if is_pdf else os.path.splitext(filename)[1] or ".png"
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        uploaded.save(tmp.name)
        tmp_path = tmp.name

    try:
        from PIL import Image as PILImage
        if is_pdf:
            from pdf2image import convert_from_path
            images = convert_from_path(tmp_path, dpi=200)
        elif is_image:
            images = [PILImage.open(tmp_path).convert("RGB")]
        else:
            return jsonify({"error": "รองรับเฉพาะไฟล์ PDF, PNG, JPG, JPEG, WEBP เท่านั้น"}), 400
    except Exception as e:
        os.unlink(tmp_path)
        return jsonify({"error": str(e)}), 500

    def generate():
        try:
            yield from _extract_stream_ndjson(images, api_key, filename)
        finally:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass

    return Response(
        stream_with_context(generate()),
        mimetype="application/x-ndjson",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


_CLAUDE_SYSTEM_PROMPT = (
    "You are a data extraction assistant for a travel agency's internal pricing system. "
    "Your job is to read tour operator rate sheets / price contracts and extract structured pricing data. "
    "Always respond with valid JSON only — no markdown, no explanations."
)

_CLAUDE_USER_PROMPT = """Extract ALL pricing data from this tour operator contract/rate sheet image(s).
CRITICAL: You MUST extract EVERY SINGLE ROW that contains a price or rate — do NOT skip, summarize, or truncate any row.

Return ONLY a JSON object in this exact format (no markdown code blocks, no extra text):
{
  "company_name": "name of the tour operator / supplier company",
  "items": [
    {
      "product_name": "Oneday Tour Type | Program Name | From → To | Passenger Type",
      "departure_time": "DEP. 08:00 - ARR. 17:00",
      "net_rate": 1000,
      "selling_rate": 1500,
      "notes": "any other remarks or conditions"
    }
  ]
}

Rules:
- company_name: the operator/supplier name (not the travel agent)
- product_name: combine ALL available fields below, separated by " | " — include every piece of information found:
    1. SECTION / BOAT TYPE header — CRITICAL: Many contracts have multiple sections separated by bold headers such as:
       "Big Boat", "Speed Boat", "Speedboat", "Big Boat Program", "Speed Boat Program",
       "เรือใหญ่", "เรือเร็ว", "Joint Tour", "Private Tour", "Standard", "Premium",
       "Oneday Tour", "Liveaboard", "Transfer", "Package", or any other section divider.
       You MUST include this header as the FIRST segment of product_name for EVERY row that falls under that section.
       NEVER drop or omit the section/boat type header — it is what distinguishes one group of products from another.
    2. Program name (e.g. "Phi Phi Island Tour", "Similan Diving Day Trip", "ATV Adventure") — always include
    3. Route: departure point → destination (e.g. "Phuket → Phi Phi", "Khao Lak → Similan", "Krabi → Railay") — look for columns: From/To, From, Route, Departure Point, Origin/Destination — include if found
    4. PASSENGER TYPE / AGE RANGE — ALWAYS append as the LAST segment of product_name whenever the row has a specific passenger type or age category:
       - Use the exact label from the document: "Adult", "Child", "Infant", "CHD", "INF", "Pax 1-4", "Min 15 Pax", "Per Person", etc.
       - This is MANDATORY: every row that corresponds to a specific age group or pax tier MUST have that label at the end of product_name.
       - Do NOT put passenger type / age category in the notes field.
    Build product_name by joining only the fields that are actually present in the document.
    Examples (showing multi-section contract with Big Boat and Speed Boat sections):
      "Big Boat | Phi Phi Island Tour | Phuket → Phi Phi | Adult"
      "Big Boat | Phi Phi Island Tour | Phuket → Phi Phi | Child"
      "Big Boat | Phi Phi Island Tour | Phuket → Phi Phi | Infant"
      "Speed Boat | Phi Phi Island Tour | Phuket → Phi Phi | Adult"
      "Speed Boat | Phi Phi Island Tour | Phuket → Phi Phi | Child"
      "Speed Boat | Similan Day Trip | Khao Lak → Similan | Adult"
      "Liveaboard | Similan Islands | Khao Lak → Similan | Adult"
      "Transfer | Phuket Airport → Patong Hotel | 1-4 Pax"
    Include as much detail as possible — do NOT omit section header, route, or passenger type if they appear anywhere in the row or surrounding headers.
- departure_time: extract the departure/arrival time as a SEPARATE field — do NOT embed it in product_name.
    Look for columns: Time DEP.-ARR., Time, Schedule, Departure Time, Pick-up Time, เวลา, ออกเดินทาง, เวลารับ
    Format: "DEP. HH:MM - ARR. HH:MM" — use the exact values from the document.
    If only departure time is shown (no arrival), use: "DEP. HH:MM"
    If the row says Full Day / Half Day / duration only (e.g. "Full Day", "Half Day AM", "2D1N"), put that in departure_time.
    CRITICAL: If the same tour/program has MULTIPLE departure times with DIFFERENT prices, each time MUST be a SEPARATE item with its own departure_time value.
    TRANSFER TABLES: if the table has a "Pick-up Time" column PLUS a boat departure time in the column header (e.g. header says "BIG BOAT — Dep. 8:45" and the row shows pickup 07:00-07:15), combine them as: "PICKUP 07:00-07:15 | DEP. 08:45".
    If no time is found for a row, use "" (empty string).
- net_rate: agent/net cost price in THB (number only). Look for: Net Rate, Net Price, Agent Rate, Net, Cost
- selling_rate: retail/public price in THB. Look for: Selling Rate, Public Rate, Rack Rate, Adult Rate, Full Price. Use 0 if not found.
- notes: CRITICAL — the notes field is the ONE AND ONLY place where Remark / Note / Condition text belongs. It MUST contain the COMPLETE, FULL, VERBATIM text from every remark/note/condition source on the page. Do NOT summarize, shorten, paraphrase, or omit any part. Do NOT place remark text in product_name, departure_time, or any other field.
    ABSOLUTE RULE: If the document has text labeled "Remark", "Remarks", "REMARK", "หมายเหตุ", "หมายเหตุ:", "Note", "Notes", "Condition", "Conditions", "Terms", "T&C", "Special Condition", "เงื่อนไข", footnote markers (*, **, ¹, ²), or any other annotation block — ALL of that text MUST appear in the notes field. Nothing should be dropped.
    Extract and combine ALL of the following sources into notes (joined with " | "):
    1. PAGE-LEVEL REMARK (MOST IMPORTANT): If the page has ANY general Remark / Note / หมายเหตุ / Condition section — whether at the bottom, top, side, a footnote block, a boxed callout, or free-floating text below the rate table — that is NOT tied to a specific row, copy that ENTIRE TEXT into the notes of EVERY SINGLE product on that page. Every product must carry the page-level remark. NEVER leave notes blank while page-level remarks exist.
    2. ROW-LEVEL REMARK: Also copy the ENTIRE TEXT from any per-row column named "Remark", "Remarks", "หมายเหตุ", "Note", "Notes", "Condition", "Conditions", "Remark/Note" — paste every word exactly as written, including numbers, symbols, bullet points, and line breaks (use a single space to join multiple lines).
    3. FOOTNOTES: If the row or table uses a footnote marker (e.g. "*", "**", "¹") that references text elsewhere on the page, copy that referenced footnote text into the notes of every row that carries the marker.
    4. "Extra Transfer" / "Extra Transfer Fee": if the document has a section, row, or column for extra transfer cost/conditions relating to a product, copy that full value into the notes of the matching product row.
    5. Any other special conditions NOT related to passenger type/age group (e.g. "Include transfer", "Min 2 pax", "Seasonal surcharge applies", "Valid Nov-Apr", "Not valid on public holidays", "Surcharge Peak 15 Dec–15 Jan +500", "Black-out dates").
    6. Operating dates / validity period text (e.g. "Valid 1 Nov 2025 – 31 Oct 2026", "Seasonal: Nov–Apr only") — include verbatim in every product's notes on that page.
    IMPORTANT: The notes field is allowed — and expected — to be very long. Preserve the full text; do NOT truncate. Every product on a page that has ANY remark content somewhere on the page MUST have that remark content in its notes field.
- MUST include ALL line items without exception: Adult, Child, Infant, every pax count variation, every category
- If prices vary by passenger type or group size, each must be a SEPARATE item — with the type/tier appended to product_name
- If a table header applies to multiple rows below it, repeat the header info in each row's product_name
- Do NOT stop early — extract until the LAST row of data on the page

- CRITICAL — MATRIX / CROSS-TABULATED TABLES (transfer pages, per-area rate sheets, etc.):
    Some pages have a GRID where ROWS = one dimension (e.g. pickup area, region, hotel zone) and COLUMNS = another dimension (e.g. boat type, transfer method, vehicle type, season).
    You MUST emit ONE SEPARATE item for EVERY non-empty CELL in the grid — i.e. rows × columns = total items.
    Common signals this is a matrix table:
      - Multiple price columns with different headers like "BIG BOAT / SPEED BOAT", "JOIN TRANSFER / PRIVATE VAN", "Adult / Child / Infant", "Low / High / Peak Season"
      - A header row with sub-headers (e.g. "BIG BOAT — Dep. 8:45" above one column, "SPEED BOAT — Dep. 9:30" above another)
      - An "Area" or "From/To" column listing multiple pickup/dropoff zones
    How to expand:
      - For each area row × each price column, create ONE item
      - product_name MUST include: column header (boat type / transfer method / vehicle type) + "Transfer" (or the section title) + the area text + the pricing unit (e.g. "Per Person", "Per Van", "1-4 Pax")
      - departure_time MUST combine the row's pickup time with the column's boat/service departure time: "PICKUP 07:00-07:15 | DEP. 08:45"
      - If the pickup cell contains descriptive text instead of a time (e.g. "According to boat arrival time", "By request", "Upon confirmation"), use that text VERBATIM as the pickup value: "PICKUP According to boat arrival time | DEP. 08:45"
      - If a single descriptive pickup cell SPANS multiple boat columns (merged cell), apply the same pickup text to every boat type — emit one item per boat × transfer method
      - net_rate = the number in that specific PRICE cell (Join Transfer or Private Van column). If the PRICE cell shows "NO SERVICE", "-", "N/A", or is blank → SKIP that specific cell (do NOT emit an item for it). Do NOT treat pickup-time cells as price cells.
      - notes: include the area description, pricing unit ("PRICE/PERSON/WAY", "PRICE/VAN/WAY"), max capacity ("Max 10 person/van"), and any page-level remarks
    Example — for a row "Rawai/Naiharn: Big Boat pickup 7:00, Speed Boat pickup 7:30, Join Transfer 200/pax, Private Van 800/way":
      → "Transfer | Big Boat | Rawai/Naiharn → Sea Angel Pier | Join Transfer — Per Person" with departure_time "PICKUP 7:00 | DEP. 08:45" net_rate 200
      → "Transfer | Big Boat | Rawai/Naiharn → Sea Angel Pier | Private Van (Max 10 pax) — Per Way" with departure_time "PICKUP 7:00 | DEP. 08:45" net_rate 800
      → "Transfer | Speed Boat | Rawai/Naiharn → Sea Angel Pier | Join Transfer — Per Person" with departure_time "PICKUP 7:30 | DEP. 09:30" net_rate 200
      → "Transfer | Speed Boat | Rawai/Naiharn → Sea Angel Pier | Private Van (Max 10 pax) — Per Way" with departure_time "PICKUP 7:30 | DEP. 09:30" net_rate 800
    Do NOT collapse or merge cells. Every non-empty priced cell = one item.

- Return ONLY the JSON object"""


def _extract_stream_ndjson(images, api_key, source_filename):
    """Yields NDJSON lines for a progress-aware extraction.

    Events:
      {"event":"start","pages":N,"filename":"..."}
      {"event":"page","page":i,"total":N}
      {"event":"items","page":i,"items":[...],"company_name":"..."}
      {"event":"done","company_name":"...","items":[...]}  # final dedup'd list
      {"event":"error","error":"..."}
    """
    import traceback as _tb

    def _emit(obj):
        return json.dumps(obj, ensure_ascii=False) + "\n"

    all_items = []
    company_name = ""
    try:
        yield _emit({"event": "start", "pages": len(images), "filename": source_filename})

        if not api_key:
            # OCR path is non-streaming — run once, emit result
            result = extract_with_ocr(images)
            for it in result.get("items", []):
                all_items.append(it)
            company_name = result.get("company_name", "")
            yield _emit({"event": "items", "page": 1, "items": all_items,
                         "company_name": company_name})
        else:
            from openai import OpenAI
            client = OpenAI(api_key=api_key)

            for page_idx, img in enumerate(images, start=1):
                yield _emit({"event": "page", "page": page_idx, "total": len(images)})
                # Reuse the single-page batch path from extract_with_claude by
                # constructing the same request and parsing the same way.
                buffer = BytesIO()
                img.save(buffer, format="JPEG", quality=95)
                img_data = base64.b64encode(buffer.getvalue()).decode()
                content = [
                    {"type": "image_url",
                     "image_url": {"url": f"data:image/jpeg;base64,{img_data}",
                                   "detail": "high"}},
                    {"type": "text", "text": _CLAUDE_USER_PROMPT},
                ]
                response = client.chat.completions.create(
                    model=OPENAI_MODEL,
                    max_tokens=OPENAI_MAX_TOKENS,
                    messages=[
                        {"role": "system", "content": _CLAUDE_SYSTEM_PROMPT},
                        {"role": "user", "content": content},
                    ],
                )
                text = response.choices[0].message.content.strip()
                text = re.sub(r"^```(?:json)?\s*\n?", "", text, flags=re.MULTILINE)
                text = re.sub(r"\n?```\s*$", "", text, flags=re.MULTILINE).strip()

                page_items = []
                page_company = ""
                try:
                    parsed = json.loads(text)
                    page_items = parsed.get("items", []) or []
                    page_company = parsed.get("company_name", "") or ""
                except json.JSONDecodeError:
                    page_items, page_company = _extract_partial_items(text)

                if page_idx == 1 and not company_name:
                    company_name = page_company

                all_items.extend(page_items)
                yield _emit({"event": "items", "page": page_idx,
                             "items": page_items, "company_name": company_name})

        # Final dedup (same logic as extract_with_claude)
        seen = set()
        deduped = []
        for item in all_items:
            key = (item.get("product_name", ""), str(item.get("net_rate", "")))
            if key not in seen:
                seen.add(key)
                deduped.append(item)

        yield _emit({"event": "done", "company_name": company_name, "items": deduped})

    except Exception as e:
        log.error("stream extract failed: %s", e)
        yield _emit({"event": "error", "error": str(e), "detail": _tb.format_exc()})



def _extract_partial_items(text):
    """Extract all complete item objects from a possibly-truncated JSON response.

    When GPT hits max_tokens the JSON is cut mid-stream. This function walks the
    text character-by-character to collect every *complete* {"product_name":...}
    object that was emitted before the cut, so we never lose the data we did get.
    Returns (items_list, company_name_str).
    """
    items = []
    company_name = ""

    company_match = re.search(r'"company_name"\s*:\s*"((?:\\.|[^"\\])*)"', text)
    if company_match:
        raw = company_match.group(1)
        # Decode JSON string escapes (\\" → ")
        try:
            company_name = json.loads(f'"{raw}"')
        except Exception:
            company_name = raw

    # Locate the start of the items array
    items_key_pos = text.find('"items"')
    if items_key_pos == -1:
        return items, company_name

    bracket_pos = text.find('[', items_key_pos)
    if bracket_pos == -1:
        return items, company_name

    # Walk through characters after '[', collecting complete {...} objects at depth 1
    depth = 0
    item_start = -1
    i = bracket_pos + 1

    while i < len(text):
        c = text[i]
        if c == '"':
            # Skip over string contents to avoid counting braces inside strings
            i += 1
            while i < len(text):
                if text[i] == '\\':
                    i += 2  # skip escaped character
                    continue
                if text[i] == '"':
                    break
                i += 1
        elif c == '{':
            if depth == 0:
                item_start = i
            depth += 1
        elif c == '}':
            depth -= 1
            if depth == 0 and item_start != -1:
                try:
                    obj = json.loads(text[item_start:i + 1])
                    if 'product_name' in obj:
                        items.append(obj)
                except Exception:
                    pass
                item_start = -1
        elif c == ']' and depth == 0:
            break  # clean end of items array
        i += 1

    return items, company_name


def extract_with_claude(images, api_key):
    from openai import OpenAI

    client = OpenAI(api_key=api_key)

    system_prompt = _CLAUDE_SYSTEM_PROMPT
    user_prompt = _CLAUDE_USER_PROMPT
    _unused_placeholder_user_prompt = """Extract ALL pricing data from this tour operator contract/rate sheet image(s).
CRITICAL: You MUST extract EVERY SINGLE ROW that contains a price or rate — do NOT skip, summarize, or truncate any row.

Return ONLY a JSON object in this exact format (no markdown code blocks, no extra text):
{
  "company_name": "name of the tour operator / supplier company",
  "items": [
    {
      "product_name": "Oneday Tour Type | Program Name | From → To | Passenger Type",
      "departure_time": "DEP. 08:00 - ARR. 17:00",
      "net_rate": 1000,
      "selling_rate": 1500,
      "notes": "any other remarks or conditions"
    }
  ]
}

Rules:
- company_name: the operator/supplier name (not the travel agent)
- product_name: combine ALL available fields below, separated by " | " — include every piece of information found:
    1. SECTION / BOAT TYPE header — CRITICAL: Many contracts have multiple sections separated by bold headers such as:
       "Big Boat", "Speed Boat", "Speedboat", "Big Boat Program", "Speed Boat Program",
       "เรือใหญ่", "เรือเร็ว", "Joint Tour", "Private Tour", "Standard", "Premium",
       "Oneday Tour", "Liveaboard", "Transfer", "Package", or any other section divider.
       You MUST include this header as the FIRST segment of product_name for EVERY row that falls under that section.
       NEVER drop or omit the section/boat type header — it is what distinguishes one group of products from another.
    2. Program name (e.g. "Phi Phi Island Tour", "Similan Diving Day Trip", "ATV Adventure") — always include
    3. Route: departure point → destination (e.g. "Phuket → Phi Phi", "Khao Lak → Similan", "Krabi → Railay") — look for columns: From/To, From, Route, Departure Point, Origin/Destination — include if found
    4. PASSENGER TYPE / AGE RANGE — ALWAYS append as the LAST segment of product_name whenever the row has a specific passenger type or age category:
       - Use the exact label from the document: "Adult", "Child", "Infant", "CHD", "INF", "Pax 1-4", "Min 15 Pax", "Per Person", etc.
       - This is MANDATORY: every row that corresponds to a specific age group or pax tier MUST have that label at the end of product_name.
       - Do NOT put passenger type / age category in the notes field.
    Build product_name by joining only the fields that are actually present in the document.
    Examples (showing multi-section contract with Big Boat and Speed Boat sections):
      "Big Boat | Phi Phi Island Tour | Phuket → Phi Phi | Adult"
      "Big Boat | Phi Phi Island Tour | Phuket → Phi Phi | Child"
      "Big Boat | Phi Phi Island Tour | Phuket → Phi Phi | Infant"
      "Speed Boat | Phi Phi Island Tour | Phuket → Phi Phi | Adult"
      "Speed Boat | Phi Phi Island Tour | Phuket → Phi Phi | Child"
      "Speed Boat | Similan Day Trip | Khao Lak → Similan | Adult"
      "Liveaboard | Similan Islands | Khao Lak → Similan | Adult"
      "Transfer | Phuket Airport → Patong Hotel | 1-4 Pax"
    Include as much detail as possible — do NOT omit section header, route, or passenger type if they appear anywhere in the row or surrounding headers.
- departure_time: extract the departure/arrival time as a SEPARATE field — do NOT embed it in product_name.
    Look for columns: Time DEP.-ARR., Time, Schedule, Departure Time, Pick-up Time, เวลา, ออกเดินทาง, เวลารับ
    Format: "DEP. HH:MM - ARR. HH:MM" — use the exact values from the document.
    If only departure time is shown (no arrival), use: "DEP. HH:MM"
    If the row says Full Day / Half Day / duration only (e.g. "Full Day", "Half Day AM", "2D1N"), put that in departure_time.
    CRITICAL: If the same tour/program has MULTIPLE departure times with DIFFERENT prices, each time MUST be a SEPARATE item with its own departure_time value.
    TRANSFER TABLES: if the table has a "Pick-up Time" column PLUS a boat departure time in the column header (e.g. header says "BIG BOAT — Dep. 8:45" and the row shows pickup 07:00-07:15), combine them as: "PICKUP 07:00-07:15 | DEP. 08:45".
    If no time is found for a row, use "" (empty string).
- net_rate: agent/net cost price in THB (number only). Look for: Net Rate, Net Price, Agent Rate, Net, Cost
- selling_rate: retail/public price in THB. Look for: Selling Rate, Public Rate, Rack Rate, Adult Rate, Full Price. Use 0 if not found.
- notes: CRITICAL — the notes field is the ONE AND ONLY place where Remark / Note / Condition text belongs. It MUST contain the COMPLETE, FULL, VERBATIM text from every remark/note/condition source on the page. Do NOT summarize, shorten, paraphrase, or omit any part. Do NOT place remark text in product_name, departure_time, or any other field.
    ABSOLUTE RULE: If the document has text labeled "Remark", "Remarks", "REMARK", "หมายเหตุ", "หมายเหตุ:", "Note", "Notes", "Condition", "Conditions", "Terms", "T&C", "Special Condition", "เงื่อนไข", footnote markers (*, **, ¹, ²), or any other annotation block — ALL of that text MUST appear in the notes field. Nothing should be dropped.
    Extract and combine ALL of the following sources into notes (joined with " | "):
    1. PAGE-LEVEL REMARK (MOST IMPORTANT): If the page has ANY general Remark / Note / หมายเหตุ / Condition section — whether at the bottom, top, side, a footnote block, a boxed callout, or free-floating text below the rate table — that is NOT tied to a specific row, copy that ENTIRE TEXT into the notes of EVERY SINGLE product on that page. Every product must carry the page-level remark. NEVER leave notes blank while page-level remarks exist.
    2. ROW-LEVEL REMARK: Also copy the ENTIRE TEXT from any per-row column named "Remark", "Remarks", "หมายเหตุ", "Note", "Notes", "Condition", "Conditions", "Remark/Note" — paste every word exactly as written, including numbers, symbols, bullet points, and line breaks (use a single space to join multiple lines).
    3. FOOTNOTES: If the row or table uses a footnote marker (e.g. "*", "**", "¹") that references text elsewhere on the page, copy that referenced footnote text into the notes of every row that carries the marker.
    4. "Extra Transfer" / "Extra Transfer Fee": if the document has a section, row, or column for extra transfer cost/conditions relating to a product, copy that full value into the notes of the matching product row.
    5. Any other special conditions NOT related to passenger type/age group (e.g. "Include transfer", "Min 2 pax", "Seasonal surcharge applies", "Valid Nov-Apr", "Not valid on public holidays", "Surcharge Peak 15 Dec–15 Jan +500", "Black-out dates").
    6. Operating dates / validity period text (e.g. "Valid 1 Nov 2025 – 31 Oct 2026", "Seasonal: Nov–Apr only") — include verbatim in every product's notes on that page.
    IMPORTANT: The notes field is allowed — and expected — to be very long. Preserve the full text; do NOT truncate. Every product on a page that has ANY remark content somewhere on the page MUST have that remark content in its notes field.
- MUST include ALL line items without exception: Adult, Child, Infant, every pax count variation, every category
- If prices vary by passenger type or group size, each must be a SEPARATE item — with the type/tier appended to product_name
- If a table header applies to multiple rows below it, repeat the header info in each row's product_name
- Do NOT stop early — extract until the LAST row of data on the page

- CRITICAL — MATRIX / CROSS-TABULATED TABLES (transfer pages, per-area rate sheets, etc.):
    Some pages have a GRID where ROWS = one dimension (e.g. pickup area, region, hotel zone) and COLUMNS = another dimension (e.g. boat type, transfer method, vehicle type, season).
    You MUST emit ONE SEPARATE item for EVERY non-empty CELL in the grid — i.e. rows × columns = total items.
    Common signals this is a matrix table:
      - Multiple price columns with different headers like "BIG BOAT / SPEED BOAT", "JOIN TRANSFER / PRIVATE VAN", "Adult / Child / Infant", "Low / High / Peak Season"
      - A header row with sub-headers (e.g. "BIG BOAT — Dep. 8:45" above one column, "SPEED BOAT — Dep. 9:30" above another)
      - An "Area" or "From/To" column listing multiple pickup/dropoff zones
    How to expand:
      - For each area row × each price column, create ONE item
      - product_name MUST include: column header (boat type / transfer method / vehicle type) + "Transfer" (or the section title) + the area text + the pricing unit (e.g. "Per Person", "Per Van", "1-4 Pax")
      - departure_time MUST combine the row's pickup time with the column's boat/service departure time: "PICKUP 07:00-07:15 | DEP. 08:45"
      - If the pickup cell contains descriptive text instead of a time (e.g. "According to boat arrival time", "By request", "Upon confirmation"), use that text VERBATIM as the pickup value: "PICKUP According to boat arrival time | DEP. 08:45"
      - If a single descriptive pickup cell SPANS multiple boat columns (merged cell), apply the same pickup text to every boat type — emit one item per boat × transfer method
      - net_rate = the number in that specific PRICE cell (Join Transfer or Private Van column). If the PRICE cell shows "NO SERVICE", "-", "N/A", or is blank → SKIP that specific cell (do NOT emit an item for it). Do NOT treat pickup-time cells as price cells.
      - notes: include the area description, pricing unit ("PRICE/PERSON/WAY", "PRICE/VAN/WAY"), max capacity ("Max 10 person/van"), and any page-level remarks
    Example — for a row "Rawai/Naiharn: Big Boat pickup 7:00, Speed Boat pickup 7:30, Join Transfer 200/pax, Private Van 800/way":
      → "Transfer | Big Boat | Rawai/Naiharn → Sea Angel Pier | Join Transfer — Per Person" with departure_time "PICKUP 7:00 | DEP. 08:45" net_rate 200
      → "Transfer | Big Boat | Rawai/Naiharn → Sea Angel Pier | Private Van (Max 10 pax) — Per Way" with departure_time "PICKUP 7:00 | DEP. 08:45" net_rate 800
      → "Transfer | Speed Boat | Rawai/Naiharn → Sea Angel Pier | Join Transfer — Per Person" with departure_time "PICKUP 7:30 | DEP. 09:30" net_rate 200
      → "Transfer | Speed Boat | Rawai/Naiharn → Sea Angel Pier | Private Van (Max 10 pax) — Per Way" with departure_time "PICKUP 7:30 | DEP. 09:30" net_rate 800
    Do NOT collapse or merge cells. Every non-empty priced cell = one item.

- Return ONLY the JSON object"""

    # Process 1 page per batch — maximum token budget per page for complete extraction
    all_items = []
    company_name = ""
    page_batches = [images[i:i+1] for i in range(0, len(images), 1)]

    for batch_idx, batch in enumerate(page_batches):
        content = []
        for img in batch:
            buffer = BytesIO()
            img.save(buffer, format="JPEG", quality=95)
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
            model=OPENAI_MODEL,
            max_tokens=OPENAI_MAX_TOKENS,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": content}
            ]
        )

        text = response.choices[0].message.content.strip()
        finish_reason = response.choices[0].finish_reason
        log.info(f"[GPT] batch={batch_idx} finish_reason={finish_reason!r} response_len={len(text)} preview={text[:120]!r}")
        # Detect refusal
        refusal_phrases = ["i'm sorry", "i cannot", "i can't", "unable to assist", "can't assist", "cannot assist"]
        if any(p in text.lower() for p in refusal_phrases) and "{" not in text:
            log.warning(f"[GPT] REFUSAL detected, skipping batch {batch_idx}")
            continue

        # Strip markdown code fences
        text = re.sub(r"^```(?:json)?\s*\n?", "", text, flags=re.MULTILINE)
        text = re.sub(r"\n?```\s*$", "", text, flags=re.MULTILINE).strip()

        # If truncated (finish_reason='length'), request continuation (loop up to 5 times)
        if finish_reason == "length":
            log.info(f"[GPT] batch={batch_idx} TRUNCATED — requesting continuation loop")
            partial_items, partial_company = _extract_partial_items(text)
            log.info(f"[GPT] Partial extraction before continuation: {len(partial_items)} items")
            all_cont_items = list(partial_items)
            conversation = [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": content},
                {"role": "assistant", "content": text},
            ]

            for cont_round in range(OPENAI_CONTINUATION_ROUNDS):
                last_item = all_cont_items[-1] if all_cont_items else {}
                continuation_prompt = (
                    f"Your previous response was cut off. You have extracted {len(all_cont_items)} items so far. "
                    f"The last item extracted was product_name={last_item.get('product_name','(none)')!r}. "
                    f"Continue from AFTER that item — list ALL REMAINING rows you have NOT yet output. "
                    f"Return ONLY: {{\"items\": [...remaining items only...]}}. "
                    f"Do NOT repeat any item already listed. If no more items remain, return {{\"items\": []}}."
                )
                try:
                    conversation.append({"role": "user", "content": continuation_prompt})
                    cont_response = client.chat.completions.create(
                        model=OPENAI_MODEL,
                        max_tokens=OPENAI_MAX_TOKENS,
                        messages=conversation
                    )
                    cont_text = cont_response.choices[0].message.content.strip()
                    cont_finish = cont_response.choices[0].finish_reason
                    cont_text = re.sub(r"^```(?:json)?\s*\n?", "", cont_text, flags=re.MULTILINE)
                    cont_text = re.sub(r"\n?```\s*$", "", cont_text, flags=re.MULTILINE).strip()
                    log.info(f"[GPT] Continuation round={cont_round} finish={cont_finish!r} len={len(cont_text)}")
                    conversation.append({"role": "assistant", "content": cont_text})
                    try:
                        cont_result = json.loads(cont_text)
                        cont_items = cont_result.get("items", [])
                    except Exception:
                        cont_items, _ = _extract_partial_items(cont_text)
                    log.info(f"[GPT] Continuation round={cont_round} items: {len(cont_items)}")
                    all_cont_items.extend(cont_items)
                    if cont_finish != "length" or not cont_items:
                        break  # done — no more truncation or no new items
                except Exception as ce:
                    log.warning(f"[GPT] Continuation round={cont_round} failed: {ce}")
                    break

            if batch_idx == 0 and not company_name:
                company_name = partial_company or ""
            all_items.extend(all_cont_items)
            continue

        try:
            batch_result = json.loads(text)
            log.info(f"[GPT] JSON parse OK: {len(batch_result.get('items',[]))} items")
        except json.JSONDecodeError as je:
            log.warning(f"[GPT] JSON parse FAILED: {je} — trying partial extraction, tail={text[-100:]!r}")
            partial_items, partial_company = _extract_partial_items(text)
            if partial_items:
                log.info(f"[GPT] Partial extraction rescued {len(partial_items)} items")
                if batch_idx == 0 and not company_name:
                    company_name = partial_company or ""
                all_items.extend(partial_items)
            else:
                log.warning(f"[GPT] No items recovered, skipping batch {batch_idx}")
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
@require_auth
def import_sheets():
    data = request.json or {}
    items = data.get("items", [])
    company = (data.get("company_name") or "").strip()
    spreadsheet_id = (data.get("spreadsheet_id") or SPREADSHEET_ID).strip()
    overwrite = bool(data.get("overwrite", False))

    if not items:
        return jsonify({"error": "ไม่มีข้อมูลที่จะนำเข้า"}), 400
    if not isinstance(items, list):
        return jsonify({"error": "items ต้องเป็น array"}), 400
    if not _SHEET_ID_RE.match(spreadsheet_id):
        return jsonify({"error": f"spreadsheet_id ไม่ถูกต้อง: {spreadsheet_id!r}"}), 400

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

        # Sheet column layout: E=Operator, F=Product name (incl. departure time appended), G=Net Rate, H=Selling Rate,
        # I=Profit(formula), J=Profit(%), K=Margin%, L=10%Commission, M-P=other cols, Q=Notes (หมายเหตุ)
        # NOTE: use ws.update() with explicit range E:Q — NOT append_rows()
        #       because append_rows() detects the table starting at col E and offsets all data by 4 cols.

        # Read all existing data for dedup + finding next empty row
        # Map key -> list of sheet row indices (1-based) so we can overwrite in place
        existing_keys_rows = {}
        all_values = ws.get_all_values()
        for idx, row in enumerate(all_values[4:], start=5):  # sheet row 5 onward
            e = row[4].strip() if len(row) > 4 else ""   # col E = Operator / company
            f = row[5].strip() if len(row) > 5 else ""   # col F = Product Name
            k = (e + "|" + f).lower()
            if k != "|":
                existing_keys_rows.setdefault(k, []).append(idx)

        # Find the next empty row (scan from bottom, look for last row with E or F data)
        next_row = 5  # default: start at row 5 (after 4 header rows)
        for i in range(len(all_values) - 1, 3, -1):
            e = all_values[i][4].strip() if len(all_values[i]) > 4 else ""
            f = all_values[i][5].strip() if len(all_values[i]) > 5 else ""
            if e or f:
                next_row = i + 2  # i is 0-indexed; +1 for 1-indexed sheet, +1 for next row
                break

        # Classify: new rows vs duplicates
        new_rows = []           # rows to append at the bottom
        duplicates = []         # list of {name_full, row_data, existing_rows}
        for item in items:
            name = item.get("product_name", "")
            dt = (item.get("departure_time") or "").strip()
            # Append departure/pickup time to product name so it lives in column F
            name_full = f"{name} | {dt}" if dt and name else (dt or name)
            k = (company.strip() + "|" + name_full.strip()).lower()
            row_data = [
                company,        # col E = Operator name
                name_full,      # col F = Product / tour name (incl. departure time)
                item.get("net_rate", item.get("net_price", "")),                                # col G = Net Rate
                item.get("selling_rate", item.get("public_rate", item.get("cost", ""))),        # col H = Selling Rate
                "",             # col I = Profit amount (formula)
                "",             # col J = Profit% (formula)
                "",             # col K = Margin% (formula)
                "",             # col L = 10% Commission (formula)
                "",             # col M (empty)
                "",             # col N (empty)
                "",             # col O (empty)
                "",             # col P (empty)
                item.get("notes", "")  # col Q = Notes (หมายเหตุ)
            ]
            if k in existing_keys_rows:
                duplicates.append({
                    "name_full": name_full,
                    "row_data": row_data,
                    "existing_rows": existing_keys_rows[k]
                })
            else:
                new_rows.append(row_data)

        # If duplicates exist and user hasn't confirmed overwrite, return conflict
        if duplicates and not overwrite:
            return jsonify({
                "success": False,
                "conflict": True,
                "duplicates": [d["name_full"] for d in duplicates],
                "new_count": len(new_rows),
                "message": f"พบรายการซ้ำ {len(duplicates)} รายการ ต้องการลบของเดิมและลงทับไหม?"
            })

        # If overwrite=True, update duplicates in place (keeping their row positions)
        overwritten = 0
        if overwrite and duplicates:
            for dup in duplicates:
                existing_rows = sorted(dup["existing_rows"])
                target_row = existing_rows[0]
                # Update first matching row with new data
                ws.update(
                    f"E{target_row}:Q{target_row}",
                    [dup["row_data"]],
                    value_input_option="USER_ENTERED"
                )
                overwritten += 1
                # If key matched multiple rows, clear the duplicates beyond the first
                for extra in existing_rows[1:]:
                    ws.batch_clear([f"E{extra}:Q{extra}"])

        # Append new (non-duplicate) rows at the bottom
        if new_rows:
            end_row = next_row + len(new_rows) - 1
            ws.update(f"E{next_row}:Q{end_row}", new_rows, value_input_option="USER_ENTERED")

        parts = []
        if new_rows:
            parts.append(f"เพิ่ม {len(new_rows)} รายการ")
        if overwritten:
            parts.append(f"ลงทับ {overwritten} รายการ")
        summary = " / ".join(parts) if parts else "ไม่มีข้อมูลใหม่"

        return jsonify({
            "success": True,
            "rows_added": len(new_rows),
            "rows_overwritten": overwritten,
            "message": f"✅ นำเข้าสำเร็จ: {summary}"
        })

    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500


# ─── Start ────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
