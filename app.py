#!/usr/bin/env python3
"""
Contract Data Importer Ã¢ÂÂ Backend (Railway)
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

# CORS Ã¢ÂÂ Ã Â¸Â­Ã Â¸ÂÃ Â¸Â¸Ã Â¸ÂÃ Â¸Â²Ã Â¸Â frontend Vercel Ã Â¹ÂÃ Â¸Â£Ã Â¸ÂµÃ Â¸Â¢Ã Â¸Â API Ã Â¹ÂÃ Â¸ÂÃ Â¹Â
CORS(app, origins="*")

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1KWqJVYfoaRg3DwslW2zSQmPgScPbE9Z-0v-Ijwtdpms")
SHEET_GID = int(os.environ.get("SHEET_GID", "384942453"))


# Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂ Health Check Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

@app.route("/")
def health():
    return jsonify({
        "status": "ok",
        "service": "Contract Importer API Ã¢ÂÂ Love Andaman",
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


# Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂ Extract Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

@app.route("/api/extract", methods=["POST"])
def extract():
    if "pdf" not in request.files:
        return jsonify({"error": "Ã Â¹ÂÃ Â¸Â¡Ã Â¹ÂÃ Â¸ÂÃ Â¸ÂÃ Â¹ÂÃ Â¸ÂÃ Â¸Â¥Ã Â¹Â PDF"}), 400

    pdf_file = request.files["pdf"]
    api_key = os.environ.get("OPENAI_API_KEY", "")

    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        pdf_file.save(tmp.name)
        tmp_path = tmp.name

    try:
        from pdf2image import convert_from_path
        images = convert_from_path(tmp_path, dpi=150)

        if api_key:
            result = extract_with_openai(images, api_key)
        else:
            result = extract_with_ocr(images)

        return jsonify(result)

    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        os.unlink(tmp_path)


def extract_with_openai(images, api_key):
    from openai import OpenAI

    client = OpenAI(api_key=api_key)
    content = []

    for img in images[:4]:
        buffer = BytesIO()
        img.save(buffer, format="PNG")
        img_data = base64.b64encode(buffer.getvalue()).decode()
        content.append({
            "type": "image_url",
            "image_url": {
                "url": f"data:image/png;base64,{img_data}"
            }
        })
    content.append({"type": "text", "text": """


Ã Â¹ÂÃ Â¸ÂÃ Â¸ÂÃ Â¸ÂµÃ Â¹ÂÃ Â¹ÂÃ Â¸Â¥Ã Â¸Â°Ã Â¸ÂÃ Â¸Â­Ã Â¸ÂÃ Â¸ÂÃ Â¸Â¥Ã Â¸Â±Ã Â¸ÂÃ Â¹ÂÃ Â¸ÂÃ Â¹ÂÃ Â¸Â JSON Ã Â¹ÂÃ Â¸ÂÃ Â¹ÂÃ Â¸Â²Ã Â¸ÂÃ Â¸Â±Ã Â¹ÂÃ Â¸Â Ã Â¹ÂÃ Â¸Â¡Ã Â¹ÂÃ Â¸Â¡Ã Â¸ÂµÃ Â¸ÂÃ Â¹ÂÃ Â¸Â­Ã Â¸ÂÃ Â¸Â§Ã Â¸Â²Ã Â¸Â¡Ã Â¸Â­Ã Â¸Â·Ã Â¹ÂÃ Â¸Â:

{
  "company_name": "Ã Â¸ÂÃ Â¸Â·Ã Â¹ÂÃ Â¸Â­Ã Â¸ÂÃ Â¸Â£Ã Â¸Â´Ã Â¸Â©Ã Â¸Â±Ã Â¸ÂÃ Â¸ÂÃ Â¸Â¹Ã Â¹ÂÃ Â¹ÂÃ Â¸Â«Ã Â¹ÂÃ Â¸ÂÃ Â¸Â£Ã Â¸Â´Ã Â¸ÂÃ Â¸Â²Ã Â¸Â£Ã Â¸ÂÃ Â¸Â±Ã Â¸Â§Ã Â¸Â£Ã Â¹Â (operator/supplier Ã Â¹ÂÃ Â¸Â¡Ã Â¹ÂÃ Â¹ÂÃ Â¸ÂÃ Â¹Â travel agent)",
  "items": [
    {
      "product_name": "Ã Â¸ÂÃ Â¸Â·Ã Â¹ÂÃ Â¸Â­Ã Â¹ÂÃ Â¸ÂÃ Â¸Â£Ã Â¹ÂÃ Â¸ÂÃ Â¸Â£Ã Â¸Â¡Ã Â¸ÂÃ Â¸Â±Ã Â¸Â§Ã Â¸Â£Ã Â¹Â/Ã Â¸ÂªÃ Â¸Â´Ã Â¸ÂÃ Â¸ÂÃ Â¹ÂÃ Â¸Â²",
      "net_price": 1000,
      "cost": 1000,
      "notes": "Ã Â¸Â«Ã Â¸Â¡Ã Â¸Â²Ã Â¸Â¢Ã Â¹ÂÃ Â¸Â«Ã Â¸ÂÃ Â¸Â¸ Ã Â¹ÂÃ Â¸ÂÃ Â¹ÂÃ Â¸Â Adult/Child, Ã Â¸ÂÃ Â¸Â³Ã Â¸ÂÃ Â¸Â§Ã Â¸ÂÃ Â¸ÂÃ Â¸Â, Ã Â¸Â«Ã Â¸Â¡Ã Â¸Â§Ã Â¸ÂÃ Â¸Â«Ã Â¸Â¡Ã Â¸Â¹Ã Â¹Â"
    }
  ]
}

Ã Â¸ÂÃ Â¸ÂÃ Â¸ÂÃ Â¸Â²Ã Â¸Â£Ã Â¸ÂÃ Â¸Â¶Ã Â¸ÂÃ Â¸ÂÃ Â¹ÂÃ Â¸Â­Ã Â¸Â¡Ã Â¸Â¹Ã Â¸Â¥:
- company_name: Ã Â¸ÂÃ Â¸Â£Ã Â¸Â´Ã Â¸Â©Ã Â¸Â±Ã Â¸ÂÃ Â¸ÂÃ Â¸Â¹Ã Â¹ÂÃ Â¹ÂÃ Â¸Â«Ã Â¹ÂÃ Â¸ÂÃ Â¸Â£Ã Â¸Â´Ã Â¸ÂÃ Â¸Â²Ã Â¸Â£Ã Â¸ÂÃ Â¸Â±Ã Â¸Â§Ã Â¸Â£Ã Â¹Â (Ã Â¹ÂÃ Â¸Â¡Ã Â¹ÂÃ Â¹ÂÃ Â¸ÂÃ Â¹Â travel agent Ã Â¸Â«Ã Â¸Â£Ã Â¸Â·Ã Â¸Â­ agent Ã Â¸ÂÃ Â¸ÂµÃ Â¹ÂÃ Â¸ÂªÃ Â¹ÂÃ Â¸Â contract Ã Â¸Â¡Ã Â¸Â²)
- product_name: Ã Â¸ÂÃ Â¸Â·Ã Â¹ÂÃ Â¸Â­Ã Â¸ÂÃ Â¸Â±Ã Â¸Â§Ã Â¸Â£Ã Â¹Â/Ã Â¹ÂÃ Â¸ÂÃ Â¸Â£Ã Â¹ÂÃ Â¸ÂÃ Â¸Â£Ã Â¸Â¡Ã Â¹ÂÃ Â¸ÂÃ Â¹ÂÃ Â¸Â¥Ã Â¸Â°Ã Â¸Â£Ã Â¸Â²Ã Â¸Â¢Ã Â¸ÂÃ Â¸Â²Ã Â¸Â£Ã Â¸Â­Ã Â¸Â¢Ã Â¹ÂÃ Â¸Â²Ã Â¸ÂÃ Â¸ÂÃ Â¸Â±Ã Â¸ÂÃ Â¹ÂÃ Â¸ÂÃ Â¸Â
- net_price: Ã Â¸Â£Ã Â¸Â²Ã Â¸ÂÃ Â¸Â² NET Ã Â¸ÂÃ Â¸ÂµÃ Â¹Â agent Ã Â¸ÂÃ Â¹ÂÃ Â¸Â²Ã Â¸Â¢ (Ã Â¸ÂÃ Â¸Â±Ã Â¸Â§Ã Â¹ÂÃ Â¸Â¥Ã Â¸Â THB Ã Â¸Â­Ã Â¸Â¢Ã Â¹ÂÃ Â¸Â²Ã Â¸ÂÃ Â¹ÂÃ Â¸ÂÃ Â¸ÂµÃ Â¸Â¢Ã Â¸Â§ Ã Â¹ÂÃ Â¸Â¡Ã Â¹ÂÃ Â¸Â¡Ã Â¸ÂµÃ Â¸Â«Ã Â¸ÂÃ Â¹ÂÃ Â¸Â§Ã Â¸Â¢)
- cost: Ã Â¹ÂÃ Â¸Â«Ã Â¸Â¡Ã Â¸Â·Ã Â¸Â­Ã Â¸Â net_price Ã Â¸Â«Ã Â¸Â²Ã Â¸ÂÃ Â¹ÂÃ Â¸Â¡Ã Â¹ÂÃ Â¸Â¡Ã Â¸Âµ cost column Ã Â¹ÂÃ Â¸Â¢Ã Â¸ÂÃ Â¸ÂÃ Â¹ÂÃ Â¸Â²Ã Â¸ÂÃ Â¸Â«Ã Â¸Â²Ã Â¸Â
- Ã Â¸Â£Ã Â¸Â§Ã Â¸Â¡Ã Â¸ÂÃ Â¸Â¸Ã Â¸ÂÃ Â¸ÂªÃ Â¸Â´Ã Â¸ÂÃ Â¸ÂÃ Â¹ÂÃ Â¸Â²/Ã Â¹ÂÃ Â¸ÂÃ Â¸Â£Ã Â¹ÂÃ Â¸ÂÃ Â¸Â£Ã Â¸Â¡Ã Â¸ÂÃ Â¸ÂµÃ Â¹ÂÃ Â¸Â¡Ã Â¸ÂµÃ Â¹ÂÃ Â¸ÂÃ Â¹ÂÃ Â¸Â­Ã Â¸ÂÃ Â¸ÂªÃ Â¸Â²Ã Â¸Â£ (Adult, Child, Infant, Ã Â¸ÂÃ Â¹ÂÃ Â¸Â²Ã Â¸ÂÃ Â¸ÂÃ Â¸Â³Ã Â¸ÂÃ Â¸Â§Ã Â¸ÂÃ Â¸ÂÃ Â¸Â)
- Ã Â¸Â«Ã Â¸Â²Ã Â¸ÂÃ Â¸Â£Ã Â¸Â²Ã Â¸ÂÃ Â¸Â²Ã Â¹ÂÃ Â¸ÂÃ Â¸ÂÃ Â¸ÂÃ Â¹ÂÃ Â¸Â²Ã Â¸ÂÃ Â¸ÂÃ Â¸Â²Ã Â¸Â¡Ã Â¸ÂÃ Â¸Â³Ã Â¸ÂÃ Â¸Â§Ã Â¸ÂÃ Â¸ÂÃ Â¸Â Ã Â¹ÂÃ Â¸Â«Ã Â¹ÂÃ Â¹ÂÃ Â¸Â¢Ã Â¸ÂÃ Â¹ÂÃ Â¸ÂÃ Â¹ÂÃ Â¸Â items Ã Â¸ÂÃ Â¸Â£Ã Â¹ÂÃ Â¸Â­Ã Â¸Â¡Ã Â¸Â£Ã Â¸Â°Ã Â¸ÂÃ Â¸Â¸Ã Â¹ÂÃ Â¸Â notes
- Ã Â¸ÂÃ Â¸Â­Ã Â¸ÂÃ Â¸ÂÃ Â¸Â¥Ã Â¸Â±Ã Â¸ÂÃ Â¹ÂÃ Â¸ÂÃ Â¹ÂÃ Â¸Â JSON Ã Â¸Â­Ã Â¸Â¢Ã Â¹ÂÃ Â¸Â²Ã Â¸ÂÃ Â¹ÂÃ Â¸ÂÃ Â¸ÂµÃ Â¸Â¢Ã Â¸Â§ Ã Â¹ÂÃ Â¸Â¡Ã Â¹ÂÃ Â¸Â¡Ã Â¸Âµ markdown code block"""
    })

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
        raise ValueError(f"Ã Â¹ÂÃ Â¸Â¡Ã Â¹ÂÃ Â¸ÂªÃ Â¸Â²Ã Â¸Â¡Ã Â¸Â²Ã Â¸Â£Ã Â¸ÂÃ Â¹ÂÃ Â¸ÂÃ Â¸Â¥Ã Â¸Â JSON: {text[:300]}")


def extract_with_ocr(images):
    try:
        import pytesseract
    except ImportError:
        return {
            "company_name": "",
            "items": [],
            "warning": "Ã¢ÂÂ Ã¯Â¸Â Ã Â¹ÂÃ Â¸Â¡Ã Â¹ÂÃ Â¸ÂÃ Â¸Â OPENAI_API_KEY Ã¢ÂÂ Ã Â¸ÂÃ Â¸Â£Ã Â¸Â¸Ã Â¸ÂÃ Â¸Â²Ã Â¸ÂÃ Â¸Â±Ã Â¹ÂÃ Â¸ÂÃ Â¸ÂÃ Â¹ÂÃ Â¸Â² Environment Variable Ã Â¸ÂÃ Â¸Â Railway"
        }

    full_text = ""
    for img in images:
        full_text += pytesseract.image_to_string(img, lang="eng") + "\n"

    company = ""
    for pattern in [r"Operator name[:\s]+([^\n\r]+)", r"Ã Â¸ÂÃ Â¸Â£Ã Â¸Â´Ã Â¸Â©Ã Â¸Â±Ã Â¸Â[:\s]+([^\n\r]+)"]:
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
        "warning": "Ã¢ÂÂ Ã¯Â¸Â Ã Â¹ÂÃ Â¸ÂÃ Â¹Â OCR Ã Â¸ÂÃ Â¸Â£Ã Â¸Â£Ã Â¸Â¡Ã Â¸ÂÃ Â¸Â² Ã Â¸ÂÃ Â¸Â£Ã Â¸Â¸Ã Â¸ÂÃ Â¸Â²Ã Â¸ÂÃ Â¸Â£Ã Â¸Â§Ã Â¸ÂÃ Â¸ÂªÃ Â¸Â­Ã Â¸ÂÃ Â¸ÂÃ Â¹ÂÃ Â¸Â­Ã Â¸Â¡Ã Â¸Â¹Ã Â¸Â¥Ã Â¸ÂÃ Â¹ÂÃ Â¸Â­Ã Â¸ÂÃ Â¸ÂÃ Â¸Â³Ã Â¹ÂÃ Â¸ÂÃ Â¹ÂÃ Â¸Â²"
    }


# Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂ Import to Sheets Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

@app.route("/api/import-sheets", methods=["POST"])
def import_sheets():
    data = request.json or {}
    items = data.get("items", [])
    company = data.get("company_name", "")
    spreadsheet_id = data.get("spreadsheet_id", SPREADSHEET_ID)

    if not items:
        return jsonify({"error": "Ã Â¹ÂÃ Â¸Â¡Ã Â¹ÂÃ Â¸Â¡Ã Â¸ÂµÃ Â¸ÂÃ Â¹ÂÃ Â¸Â­Ã Â¸Â¡Ã Â¸Â¹Ã Â¸Â¥Ã Â¸ÂÃ Â¸ÂµÃ Â¹ÂÃ Â¸ÂÃ Â¸Â°Ã Â¸ÂÃ Â¸Â³Ã Â¹ÂÃ Â¸ÂÃ Â¹ÂÃ Â¸Â²"}), 400

    # Ã Â¸Â­Ã Â¹ÂÃ Â¸Â²Ã Â¸Â credentials Ã Â¸ÂÃ Â¸Â²Ã Â¸Â Environment Variable
    creds_json_str = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")
    if not creds_json_str:
        return jsonify({
            "error": "Ã Â¹ÂÃ Â¸Â¡Ã Â¹ÂÃ Â¸ÂÃ Â¸Â GOOGLE_CREDENTIALS_JSON",
            "help": "Ã Â¸ÂÃ Â¸Â£Ã Â¸Â¸Ã Â¸ÂÃ Â¸Â²Ã Â¸ÂÃ Â¸Â±Ã Â¹ÂÃ Â¸ÂÃ Â¸ÂÃ Â¹ÂÃ Â¸Â² Environment Variable GOOGLE_CREDENTIALS_JSON Ã Â¸ÂÃ Â¸Â Railway"
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

        # Ã Â¸Â«Ã Â¸Â² worksheet Ã Â¸ÂÃ Â¸Â²Ã Â¸Â¡ GID
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
            "message": f"Ã Â¸ÂÃ Â¸Â³Ã Â¹ÂÃ Â¸ÂÃ Â¹ÂÃ Â¸Â²Ã Â¸ÂÃ Â¹ÂÃ Â¸Â­Ã Â¸Â¡Ã Â¸Â¹Ã Â¸Â¥Ã Â¸ÂªÃ Â¸Â³Ã Â¹ÂÃ Â¸Â£Ã Â¹ÂÃ Â¸Â {len(rows)} Ã Â¸Â£Ã Â¸Â²Ã Â¸Â¢Ã Â¸ÂÃ Â¸Â²Ã Â¸Â£"
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂ Start Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
