"""security.py — Auth, rate limiting, input validation"""

import time, re, logging
from config import ADMIN_ID

logger = logging.getLogger(__name__)
_cooldowns:     dict[int, float] = {}
_unauth_count:  dict[int, int]   = {}
MAX_UNAUTH = 5

def is_admin(uid): return uid == ADMIN_ID

def check_cooldown(uid, sec):
    if uid == ADMIN_ID: return 0
    elapsed = time.time() - _cooldowns.get(uid, 0)
    return max(0, int(sec - elapsed))

def set_cooldown(uid): _cooldowns[uid] = time.time()

def track_unauth(uid):
    _unauth_count[uid] = _unauth_count.get(uid, 0) + 1
    return _unauth_count[uid]

def reset_unauth(uid): _unauth_count.pop(uid, None)

def validate_file(file_obj, max_mb):
    size = getattr(file_obj, "file_size", None)
    if size and size > max_mb * 1024 * 1024:
        return f"⚠️ File too large. Max *{max_mb} MB*."
    return None

def safe_error(e):
    msg = str(e)
    msg = re.sub(r"key=[A-Za-z0-9_\-]+", "key=***", msg)
    msg = re.sub(r"\d{8,}:[A-Za-z0-9_\-]+", "***", msg)
    msg = re.sub(r"[A-Za-z0-9_\-]{40,}", "***", msg)
    return msg
