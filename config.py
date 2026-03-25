"""
config.py — Central configuration loader
"""

import os
from dotenv import load_dotenv

load_dotenv()

# ─── Required ─────────────────────────────────────────────────────────────────
TELEGRAM_TOKEN = os.environ["TELEGRAM_TOKEN"]
ADMIN_ID        = int(os.environ["ADMIN_ID"])

# ─── Gemini API Keys (multi-key fallback) ─────────────────────────────────────
GEMINI_KEYS: list[str] = []
i = 1
while True:
    key = os.environ.get(f"GEMINI_API_KEY_{i}") or (os.environ.get("GEMINI_API_KEY") if i == 1 else None)
    if not key:
        break
    GEMINI_KEYS.append(key)
    i += 1

if not GEMINI_KEYS:
    raise ValueError("No GEMINI_API_KEY found in .env")

# ─── LlamaParse ───────────────────────────────────────────────────────────────
LLAMA_API_KEY = os.environ.get("LLAMA_API_KEY", "")   # From cloud.llamaindex.ai (free: 1000 pages/day)

# ─── Webhook ──────────────────────────────────────────────────────────────────
WEBHOOK_URL  = os.environ.get("WEBHOOK_URL", "").rstrip("/")
WEBHOOK_PORT = int(os.environ.get("WEBHOOK_PORT", 8443))
WEBHOOK_PATH = os.environ.get("WEBHOOK_PATH", "/webhook")
USE_WEBHOOK  = bool(WEBHOOK_URL)

# ─── Defaults (all overridable live from admin panel) ─────────────────────────
DEFAULT_MAX_FILE_MB    = int(os.environ.get("MAX_FILE_MB", 10))
DEFAULT_COOLDOWN_SEC   = int(os.environ.get("COOLDOWN_SECONDS", 30))
DEFAULT_DAILY_LIMIT    = int(os.environ.get("DAILY_LIMIT", 20))
DEFAULT_GEMINI_MODEL   = os.environ.get("GEMINI_MODEL", "gemini-2.0-flash")
DEFAULT_WHITELIST_ONLY = os.environ.get("WHITELIST_ONLY", "true").lower() == "true"
DEFAULT_LLAMAPARSE_ON  = os.environ.get("LLAMAPARSE_ON", "false").lower() == "true"

# ─── Paths ────────────────────────────────────────────────────────────────────
DATA_FILE = "data.json"
LOG_FILE  = "bot.log"
