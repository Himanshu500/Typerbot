"""
detector.py — Fast content type classifier using Gemini vision

One cheap API call to classify document type before full extraction.
This lets us pick the right specialized prompt for best results.
"""

import base64
import logging
import requests

logger = logging.getLogger(__name__)

CONTENT_TYPES = {
    "OFFICIAL_NOTICE":   "University notice, government order, formal notification with header/ref/date",
    "FORMAL_DOCUMENT":   "Letter, declaration, affidavit, certificate, legal document",
    "PRODUCT_LABEL":     "Product packaging, ingredient list, instructions on a physical label",
    "TEXTBOOK_SCIENCE":  "Science/chemistry/math textbook with formulas, equations, Q&A columns",
    "TEXTBOOK_GENERAL":  "General textbook, prose chapters, notes, educational content",
    "FORM_TABLE":        "Form, invoice, structured data table, receipt, spreadsheet-like",
    "MIXED":             "Combination of multiple types above",
}

DETECT_PROMPT = """Analyze this document image and classify it into EXACTLY ONE of these types:

OFFICIAL_NOTICE   - University notice, government order, formal notification with letterhead, ref number, date
FORMAL_DOCUMENT   - Letter, declaration, affidavit, certificate, contract, legal document  
PRODUCT_LABEL     - Product packaging, ingredient list, usage instructions on physical label/box
TEXTBOOK_SCIENCE  - Science/chemistry/math textbook with formulas, structures, equations, two-column Q&A
TEXTBOOK_GENERAL  - General textbook, prose chapters, educational notes, study material
FORM_TABLE        - Form, invoice, receipt, structured data, spreadsheet-like document
MIXED             - Clear combination of multiple types

Also detect:
- Primary language (English / Hindi / Mixed / Other)
- Has tables: yes/no
- Has chemical formulas: yes/no  
- Has handwriting: yes/no
- Layout: single_column / two_column / multi_column / mixed

Respond in EXACTLY this format (no other text):
TYPE: <type>
LANG: <language>
TABLES: <yes/no>
FORMULAS: <yes/no>
HANDWRITING: <yes/no>
LAYOUT: <layout>"""


def detect(file_bytes: bytes, mime_type: str, gemini_key: str, model: str = "gemini-2.0-flash") -> dict:
    """
    Classify document type with a single fast API call.
    Returns a dict with type, language, and feature flags.
    Falls back to MIXED on any error.
    """
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    payload = {
        "contents": [{
            "parts": [
                {"text": DETECT_PROMPT},
                {"inline_data": {
                    "mime_type": mime_type,
                    "data": base64.b64encode(file_bytes).decode(),
                }}
            ]
        }],
        "generationConfig": {"maxOutputTokens": 100, "temperature": 0.1},
    }

    try:
        resp = requests.post(url, params={"key": gemini_key}, json=payload, timeout=30)
        if resp.status_code != 200:
            logger.warning("Detector API error %d — defaulting to MIXED", resp.status_code)
            return _default()

        text = resp.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
        return _parse(text)

    except Exception as e:
        logger.warning("Detection failed: %s — defaulting to MIXED", e)
        return _default()


def _parse(text: str) -> dict:
    result = _default()
    for line in text.strip().splitlines():
        line = line.strip()
        if line.startswith("TYPE:"):
            val = line.split(":", 1)[1].strip().upper()
            if val in CONTENT_TYPES:
                result["type"] = val
        elif line.startswith("LANG:"):
            result["lang"] = line.split(":", 1)[1].strip()
        elif line.startswith("TABLES:"):
            result["has_tables"] = "yes" in line.lower()
        elif line.startswith("FORMULAS:"):
            result["has_formulas"] = "yes" in line.lower()
        elif line.startswith("HANDWRITING:"):
            result["has_handwriting"] = "yes" in line.lower()
        elif line.startswith("LAYOUT:"):
            result["layout"] = line.split(":", 1)[1].strip().lower()
    return result


def _default() -> dict:
    return {
        "type": "MIXED",
        "lang": "English",
        "has_tables": False,
        "has_formulas": False,
        "has_handwriting": False,
        "layout": "single_column",
    }
