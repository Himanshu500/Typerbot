"""
gemini.py — Multi-key Gemini engine with:
  - Automatic key fallback on 429/403
  - Specialized prompts per content type
  - Second styling analysis call
  - LlamaParse integration for PDFs
"""

import time
import base64
import json
import logging
import re
import requests
from datetime import datetime, timedelta
from config import GEMINI_KEYS, LLAMA_API_KEY

logger = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════════════
#  EXTRACTION PROMPTS — one per content type
# ══════════════════════════════════════════════════════════════════════════════

_BASE_RULES = """
CRITICAL RULES FOR ALL CONTENT:
- NEVER split a sentence at a physical line break — always merge wrapped lines into complete sentences
- NEVER add [Drawing of...] or [Image of...] descriptions
- Preserve ALL text exactly including numbers, dates, percentages, symbols
- Merge list items that span multiple visual lines into single complete items
- For unclear/handwritten text use [unclear: best_guess]
- Preserve original numbered format exactly (01. 02. not 1. 2.)
"""

PROMPTS = {

"OFFICIAL_NOTICE": """You are extracting content from an official notice/letter.

OUTPUT FORMAT MARKERS:
  CENTER: text          → centered text (titles, headings)
  REFDATE: ref | date   → reference number and date on same line
  BOLD: text            → bold inline text  
  SUBJECT: text         → subject line (bold)
  # text               → main heading
  ## text              → subheading
  - item               → bullet point (complete sentence)
  1. item              → numbered item (complete sentence, merge wrapped lines)
  TABLE:               → start of table
  col1 | col2 | col3   → table row
  END_TABLE            → end of table
  SIGTABLE:            → signature block (3 columns: Recommended By | Reviewed By | Approved By)
  SIGNAME: name        → name in signature block
  SIGROLE: role        → role/designation in signature block  
  COMMENTBOX: title    → bordered comment/recommendation box

RULES:
- Merge ALL wrapped sentence lines — one sentence = one paragraph  
- Bold text that appears visually bold
- Preserve superscript ordinals as: 26^th^ 1^st^ 2^nd^ 3^rd^
- Tables: extract ALL rows including header
- Signature section: always use SIGTABLE format
- If logo present: note as [LOGO: organization name]
""" + _BASE_RULES,

"FORMAL_DOCUMENT": """You are extracting content from a formal document (letter/declaration/affidavit).

OUTPUT FORMAT MARKERS:
  CENTER: text          → centered text
  # text               → title/heading
  BOLD: text            → bold inline text
  - item               → bullet point
  01. item             → numbered item (preserve original numbering format)
  TABLE:               → table start
  col1 | col2          → table row  
  END_TABLE            → table end

RULES:
- Preserve the EXACT numbered format from original (01. 02. or 1. 2. or i. ii.)
- Merge ALL wrapped lines into complete paragraphs — physical line breaks are NOT paragraph breaks
- Inline bold/italic: preserve exactly as seen
- Handwritten portions: [unclear: best_guess] or [handwritten: text]
- Keep paragraph spacing natural — blank line between sections
""" + _BASE_RULES,

"PRODUCT_LABEL": """You are extracting content from a product label/packaging.

OUTPUT FORMAT MARKERS:
  CENTER: text          → centered text (product name, title)
  # text               → main title
  BOLD: text            → bold label or term
  BOLDSTART: label      → section label that starts a paragraph (e.g. "Directions for use:")
  - item               → bullet point (ALWAYS complete sentence — merge wrapped lines)
  1. item              → numbered item (ALWAYS complete sentence — merge wrapped lines)  
  TABLE:               → table start
  col1 | col2 | col3   → table row
  END_TABLE            → table end
  SIDEBY_LEFT: text    → left column of side-by-side block
  SIDEBY_RIGHT: text   → right column of side-by-side block

RULES:
- CRITICAL: Physical line wraps in labels are NOT paragraph breaks — merge into complete sentences
- List items that wrap visually: merge ALL continuation lines into ONE list item
- Only bold text that is VISUALLY BOLD in the image — do not bold ingredient names
- Asterisk notes (*text): mark as FOOTNOTE: text
- Side-by-side blocks (Manufactured by | Factory): use SIDEBY markers
""" + _BASE_RULES,

"TEXTBOOK_SCIENCE": """You are extracting content from a science/chemistry textbook.

OUTPUT FORMAT MARKERS:
  # text               → chapter/section heading
  ## text              → subheading
  BOLDNUM: 6.4         → bold question/section number
  STRUCT:              → start of chemical structural formula block
  [formula lines]      → preserve exact spacing/alignment with spaces
  END_STRUCT           → end structural formula block
  QA_TABLE:            → start of question-answer two-column section  
  FORMULA | EXPLANATION → Q&A row (formula/structure left, explanation right)
  END_QA_TABLE         → end Q&A table
  TABLE:               → regular table
  col1 | col2          → table row
  END_TABLE            → end table
  - item               → bullet point
  1. item              → numbered item

RULES:
- Two-column layouts (formula left, explanation right): ALWAYS use QA_TABLE format
- Chemical structural formulas with bonds (|, -, =): ALWAYS wrap in STRUCT block
- Subscripts: use Unicode (₀₁₂₃₄₅₆₇₈₉) for chemical formulas
- Superscripts: use ^text^ notation (e.g. R^1^, x^2^)
- Ring structures (benzene, cyclohexane): write condensed formula (C₆H₅- or C₆H₁₁-)
- Section numbers (6.4, 6.5): use BOLDNUM marker
- Merge continuation explanatory text — do not split explanations at line breaks
""" + _BASE_RULES,

"TEXTBOOK_GENERAL": """You are extracting content from a general textbook or educational document.

OUTPUT FORMAT MARKERS:
  # text               → chapter heading
  ## text              → section heading  
  ### text             → subsection heading
  BOLD: text           → bold term or keyword
  - item               → bullet point
  1. item              → numbered item
  TABLE:               → table
  col1 | col2          → row
  END_TABLE            → end table

RULES:
- Merge ALL visual line wraps into complete paragraphs
- Preserve heading hierarchy exactly
- Bold terms/keywords inline using BOLD: marker
- Tables: full extraction with headers
""" + _BASE_RULES,

"FORM_TABLE": """You are extracting content from a form, invoice, or structured table document.

OUTPUT FORMAT MARKERS:
  # text               → document title
  FIELD: label | value → form field with label and value
  TABLE:               → table start
  col1 | col2 | col3   → table row (header row first)
  END_TABLE            → table end
  BOLD: text           → bold text
  CENTER: text         → centered text

RULES:
- Extract ALL fields including empty ones (use FIELD: label | [empty])
- Tables: preserve ALL rows and columns exactly
- Multi-line cell content: use \n within cell value
- Preserve currency symbols, units, percentages exactly
""" + _BASE_RULES,

"MIXED": """You are extracting content from a document with mixed content types.

Use the most appropriate markers from these depending on content:
  CENTER: text         → centered/title text
  # text              → main heading
  ## text             → subheading
  BOLD: text          → bold text
  BOLDSTART: label    → bold label starting a paragraph
  - item              → bullet (complete sentence)
  1. item             → numbered item (complete sentence)
  TABLE:              → table start
  col1 | col2         → row
  END_TABLE           → table end
  STRUCT:             → chemical structure block
  END_STRUCT          → end structure
  SIGTABLE:           → signature block
  REFDATE: ref | date → ref and date line
  FIELD: label | val  → form field
  FOOTNOTE: text      → footnote/asterisk note
  SIDEBY_LEFT: text   → left side-by-side block
  SIDEBY_RIGHT: text  → right side-by-side block

RULES:
- Merge ALL physical line wraps into complete sentences
- Detect and preserve all visual formatting (bold, italic, alignment)
- Tables: full extraction
- Signatures: use SIGTABLE format
""" + _BASE_RULES,
}


# ══════════════════════════════════════════════════════════════════════════════
#  STYLING PROMPT — second call for visual formatting analysis
# ══════════════════════════════════════════════════════════════════════════════

STYLE_PROMPT = """Analyze the VISUAL STYLING of this document image.
Return ONLY valid JSON, no markdown, no explanation.

{
  "doc_font_style": "sans-serif",
  "doc_base_size": "medium",
  "doc_alignment": "left",
  "line_spacing": "normal",
  "elements": [
    {
      "role": "main_title",
      "font_style": "sans-serif",
      "size_category": "xlarge",
      "bold": true,
      "italic": false,
      "align": "center",
      "spacing": "normal",
      "color": "black"
    }
  ]
}

Role options: main_title, subtitle, heading, subheading, ref_line, subject_line, body, list_item, table_header, table_cell, label, footnote, signature_name, signature_role, caption
Font style options: serif, sans-serif, mono, handwriting
Size category options: xlarge (>18pt), large (14-18pt), medium (12-14pt), normal (11-12pt), small (9-11pt), tiny (<9pt)
Align options: left, center, right, justify
Spacing options: tight, normal, loose
Color options: black, dark-blue, blue, gray, red, other

Identify ALL distinct element types present in the document."""


# ══════════════════════════════════════════════════════════════════════════════
#  KEY STATE
# ══════════════════════════════════════════════════════════════════════════════

class KeyState:
    def __init__(self, key: str, index: int):
        self.key         = key
        self.index       = index
        self.used        = 0
        self.exhausted   = False
        self.retry_after: datetime | None = None

    def is_available(self) -> bool:
        if not self.exhausted:
            return True
        if self.retry_after and datetime.now() >= self.retry_after:
            self.exhausted   = False
            self.retry_after = None
            return True
        return False

    def mark_rate_limited(self, wait_sec: int = 60):
        self.exhausted   = True
        self.retry_after = datetime.now() + timedelta(seconds=wait_sec)

    def mark_quota_exceeded(self):
        self.exhausted   = True
        tomorrow = datetime.now().replace(hour=0, minute=0, second=0) + timedelta(days=1)
        self.retry_after = tomorrow


# ══════════════════════════════════════════════════════════════════════════════
#  GEMINI ENGINE
# ══════════════════════════════════════════════════════════════════════════════

class GeminiEngine:
    def __init__(self, keys: list[str]):
        if not keys:
            raise ValueError("At least one Gemini API key required")
        self.keys  = [KeyState(k, i) for i, k in enumerate(keys)]
        self.model = "gemini-2.0-flash"

    def set_model(self, model: str):
        self.model = model

    def status(self) -> list[dict]:
        return [{
            "index":     ks.index + 1,
            "available": ks.is_available(),
            "used":      ks.used,
            "exhausted": ks.exhausted,
            "retry_after": ks.retry_after.isoformat() if ks.retry_after else None,
        } for ks in self.keys]

    def _available(self) -> list[KeyState]:
        return [ks for ks in self.keys if ks.is_available()]

    def _call(self, payload: dict, timeout: int = 60) -> str:
        """Make API call with automatic key fallback."""
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{self.model}:generateContent"

        while True:
            available = self._available()
            if not available:
                raise Exception(
                    f"⚠️ All {len(self.keys)} Gemini API key(s) are exhausted.\n"
                    "Add more keys in .env or wait for rate limit to reset."
                )
            ks = available[0]
            try:
                resp = requests.post(
                    url, params={"key": ks.key},
                    json=payload, timeout=timeout,
                )
            except requests.exceptions.Timeout:
                ks.mark_rate_limited(30)
                continue

            if resp.status_code == 200:
                ks.used += 1
                return resp.json()["candidates"][0]["content"]["parts"][0]["text"]
            elif resp.status_code == 429:
                retry = int(resp.headers.get("Retry-After", 60))
                ks.mark_rate_limited(retry)
            elif resp.status_code == 403:
                ks.mark_quota_exceeded()
            else:
                raise Exception(f"Gemini API {resp.status_code}: {resp.text[:300]}")

    def extract(self, file_bytes: bytes, mime_type: str, doc_info: dict) -> str:
        """Call 1: Extract content with type-specific prompt."""
        doc_type = doc_info.get("type", "MIXED")
        prompt   = PROMPTS.get(doc_type, PROMPTS["MIXED"])

        # Add language hint if non-English
        lang = doc_info.get("lang", "English")
        if lang not in ("English",):
            prompt += f"\n\nNOTE: Document is in {lang}. Preserve ALL text in original language exactly."

        payload = {
            "contents": [{
                "parts": [
                    {"text": prompt},
                    {"inline_data": {"mime_type": mime_type, "data": base64.b64encode(file_bytes).decode()}}
                ]
            }],
            "generationConfig": {"temperature": 0.1, "maxOutputTokens": 8192},
        }
        return self._call(payload, timeout=90)

    def analyze_style(self, file_bytes: bytes, mime_type: str) -> dict:
        """Call 2: Analyze visual styling — returns parsed style dict."""
        payload = {
            "contents": [{
                "parts": [
                    {"text": STYLE_PROMPT},
                    {"inline_data": {"mime_type": mime_type, "data": base64.b64encode(file_bytes).decode()}}
                ]
            }],
            "generationConfig": {"temperature": 0.1, "maxOutputTokens": 2048},
        }
        try:
            raw = self._call(payload, timeout=45)
            # Strip markdown code fences
            raw = re.sub(r"```(?:json)?", "", raw).strip().rstrip("`").strip()
            # Remove JS-style comments (// ...)
            raw = re.sub(r"//[^\n]*", "", raw)
            # Remove trailing commas before } or ]
            raw = re.sub(r",\s*([}\]])", r"\1", raw)
            return json.loads(raw)
        except Exception as e:
            logger.warning("Style analysis failed: %s — using defaults", e)
            return {}


# ══════════════════════════════════════════════════════════════════════════════
#  LLAMAPARSE INTEGRATION
# ══════════════════════════════════════════════════════════════════════════════

def extract_with_llamaparse(file_bytes: bytes, filename: str) -> str | None:
    """
    Parse PDF with LlamaParse (free: 1000 pages/day).
    Returns markdown string or None on failure.
    Only works for PDFs.
    """
    if not LLAMA_API_KEY:
        logger.warning("LlamaParse requested but LLAMA_API_KEY not set")
        return None

    try:
        import tempfile, os
        from llama_parse import LlamaParse
        from llama_index.core import SimpleDirectoryReader

        parser = LlamaParse(
            api_key=LLAMA_API_KEY,
            result_type="markdown",
            verbose=False,
            language="en",
        )

        # Write bytes to temp file
        suffix = ".pdf" if filename.endswith(".pdf") else ".pdf"
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
            tmp.write(file_bytes)
            tmp_path = tmp.name

        try:
            file_extractor = {suffix: parser}
            documents = SimpleDirectoryReader(
                input_files=[tmp_path],
                file_extractor=file_extractor,
            ).load_data()
            result = "\n\n".join(doc.text for doc in documents)
            logger.info("LlamaParse extracted %d chars", len(result))
            return result
        finally:
            os.unlink(tmp_path)

    except ImportError:
        logger.warning("llama-parse not installed — falling back to Gemini")
        return None
    except Exception as e:
        logger.warning("LlamaParse failed: %s — falling back to Gemini", e)
        return None


# ─── Global engine instance ───────────────────────────────────────────────────
engine = GeminiEngine(GEMINI_KEYS)
