"""
store.py — Persistent JSON data store
"""

import json
from pathlib import Path
from datetime import datetime
from config import (
    DATA_FILE, DEFAULT_MAX_FILE_MB, DEFAULT_COOLDOWN_SEC,
    DEFAULT_DAILY_LIMIT, DEFAULT_GEMINI_MODEL,
    DEFAULT_WHITELIST_ONLY, DEFAULT_LLAMAPARSE_ON,
)

_PATH = Path(DATA_FILE)


def _default() -> dict:
    return {
        "settings": {
            "max_file_mb":    DEFAULT_MAX_FILE_MB,
            "cooldown_sec":   DEFAULT_COOLDOWN_SEC,
            "daily_limit":    DEFAULT_DAILY_LIMIT,
            "gemini_model":   DEFAULT_GEMINI_MODEL,
            "whitelist_only": DEFAULT_WHITELIST_ONLY,
            "bot_enabled":    True,
            "llamaparse_on":  DEFAULT_LLAMAPARSE_ON,   # admin-toggleable
            "welcome_msg":    "👋 Hello! Send me a *PDF* or *image* and I'll extract the content into a *.docx* file.",
        },
        "users":  {},
        "stats":  {
            "total": 0, "success": 0, "failed": 0,
            "by_date": {}, "by_user": {},
            "llamaparse_used": 0,
        },
        "broadcasts": [],
        "key_stats":   {},
    }


def load() -> dict:
    if _PATH.exists():
        try:
            d = json.loads(_PATH.read_text())
            # Patch missing keys from new defaults
            d.setdefault("settings", {})
            d["settings"].setdefault("llamaparse_on", DEFAULT_LLAMAPARSE_ON)
            d["settings"].setdefault("welcome_msg", "👋 Hello! Send me a *PDF* or *image* and I'll extract the content into a *.docx* file.")
            return d
        except Exception:
            pass
    d = _default()
    save(d)
    return d


def save(d: dict):
    _PATH.write_text(json.dumps(d, indent=2, default=str))


# ─── User helpers ─────────────────────────────────────────────────────────────

def get_user(d: dict, user_id: int) -> dict | None:
    return d["users"].get(str(user_id))


def register_user(d: dict, user_id: int, full_name: str, username: str) -> bool:
    uid = str(user_id)
    if uid in d["users"]:
        d["users"][uid]["name"]     = full_name
        d["users"][uid]["username"] = username or ""
        save(d)
        return False
    d["users"][uid] = {
        "name": full_name, "username": username or "",
        "allowed": False, "blocked": False,
        "joined": datetime.now().isoformat(),
        "last_seen": datetime.now().isoformat(),
        "requests": 0, "failed_auth": 0, "note": "",
    }
    save(d)
    return True


def set_allowed(d: dict, user_id: int, allowed: bool):
    uid = str(user_id)
    if uid in d["users"]:
        d["users"][uid]["allowed"] = allowed
        save(d)


def set_blocked(d: dict, user_id: int, blocked: bool):
    uid = str(user_id)
    if uid in d["users"]:
        d["users"][uid]["blocked"] = blocked
        save(d)


def is_allowed(d: dict, user_id: int, admin_id: int) -> bool:
    if user_id == admin_id:
        return True
    u = d["users"].get(str(user_id))
    if not u:
        return False
    return u.get("allowed", False) and not u.get("blocked", False)


def pending_users(d: dict) -> list[dict]:
    return [{"id": int(uid), **u} for uid, u in d["users"].items()
            if not u.get("allowed") and not u.get("blocked")]


def all_users(d: dict) -> list[dict]:
    return [{"id": int(uid), **u} for uid, u in d["users"].items()]


# ─── Stats ────────────────────────────────────────────────────────────────────

def record_request(d: dict, user_id: int, success: bool, used_llamaparse: bool = False):
    today = datetime.now().strftime("%Y-%m-%d")
    uid   = str(user_id)
    d["stats"]["total"]   += 1
    d["stats"]["success" if success else "failed"] += 1
    d["stats"]["by_date"][today] = d["stats"]["by_date"].get(today, 0) + 1
    d["stats"]["by_user"][uid]   = d["stats"]["by_user"].get(uid, 0) + 1
    if used_llamaparse:
        d["stats"]["llamaparse_used"] = d["stats"].get("llamaparse_used", 0) + 1
    if uid in d["users"]:
        d["users"][uid]["requests"] += 1
        d["users"][uid]["last_seen"] = datetime.now().isoformat()
    save(d)


def today_count(d: dict, user_id: int) -> int:
    today = datetime.now().strftime("%Y-%m-%d")
    uid   = str(user_id)
    u     = d["users"].get(uid, {})
    if u.get("last_req_date") == today:
        return u.get("today_count", 0)
    return 0


def increment_today(d: dict, user_id: int):
    today = datetime.now().strftime("%Y-%m-%d")
    uid   = str(user_id)
    u     = d["users"].get(uid, {})
    if u.get("last_req_date") != today:
        d["users"][uid]["last_req_date"] = today
        d["users"][uid]["today_count"]   = 1
    else:
        d["users"][uid]["today_count"] = u.get("today_count", 0) + 1
    save(d)
