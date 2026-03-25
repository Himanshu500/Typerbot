"""
admin.py — Full Telegram inline admin panel
Includes LlamaParse on/off toggle and API key status
"""

import json, io, logging
from datetime import datetime
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ContextTypes
import store
from config import ADMIN_ID, LLAMA_API_KEY
from gemini import engine
from security import is_admin

logger = logging.getLogger(__name__)


def admin_only(func):
    async def wrapper(update: Update, context: ContextTypes.DEFAULT_TYPE):
        uid = update.effective_user.id
        if not is_admin(uid):
            if update.callback_query:
                await update.callback_query.answer("⛔ Unauthorized.", show_alert=True)
            else:
                await update.message.reply_text("⛔ Unauthorized.")
            return
        return await func(update, context)
    return wrapper


# ── Keyboards ──────────────────────────────────────────────────────────────────

def kb_main() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("👥 Users",        callback_data="adm:users"),
         InlineKeyboardButton("⚙️ Settings",     callback_data="adm:settings")],
        [InlineKeyboardButton("📊 Stats",        callback_data="adm:stats"),
         InlineKeyboardButton("🔑 API Keys",     callback_data="adm:keys")],
        [InlineKeyboardButton("🤖 AI Mode",      callback_data="adm:aimode"),
         InlineKeyboardButton("📢 Broadcast",    callback_data="adm:broadcast")],
        [InlineKeyboardButton("🔒 Security",     callback_data="adm:security"),
         InlineKeyboardButton("🗂️ Data",         callback_data="adm:data")],
    ])

def kb_back(to="main"):
    return InlineKeyboardMarkup([[InlineKeyboardButton("⬅️ Back", callback_data=f"adm:{to}")]])

def kb_settings(data):
    s  = data["settings"]
    on = s["bot_enabled"]
    wl = s["whitelist_only"]
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(f"🤖 Bot: {'ON ✅' if on else 'OFF 🔴'}", callback_data="adm:toggle_bot")],
        [InlineKeyboardButton(f"🔒 Whitelist: {'ON ✅' if wl else 'OFF 🔴'}", callback_data="adm:toggle_whitelist")],
        [InlineKeyboardButton(f"📦 Max File: {s['max_file_mb']} MB",    callback_data="adm:set:max_file_mb")],
        [InlineKeyboardButton(f"⏱️ Cooldown: {s['cooldown_sec']}s",     callback_data="adm:set:cooldown_sec")],
        [InlineKeyboardButton(f"📅 Daily Limit: {s['daily_limit']}",    callback_data="adm:set:daily_limit")],
        [InlineKeyboardButton("✏️ Edit Welcome Msg",                     callback_data="adm:set:welcome_msg")],
        [InlineKeyboardButton("⬅️ Back",                                 callback_data="adm:main")],
    ])

def kb_aimode(data):
    s   = data["settings"]
    lp  = s.get("llamaparse_on", False)
    mdl = s.get("gemini_model", "gemini-2.0-flash")
    has_key = bool(LLAMA_API_KEY)
    lp_label = f"📄 LlamaParse: {'ON ✅' if lp else 'OFF 🔴'}"
    if not has_key:
        lp_label += " (no key)"
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(lp_label,                                   callback_data="adm:toggle_llama")],
        [InlineKeyboardButton(f"🤖 Gemini Model: {mdl}",                  callback_data="adm:set:gemini_model")],
        [InlineKeyboardButton("📊 Key Status",                            callback_data="adm:keys")],
        [InlineKeyboardButton("ℹ️ What is LlamaParse?",                   callback_data="adm:llama_info")],
        [InlineKeyboardButton("⬅️ Back",                                  callback_data="adm:main")],
    ])

def kb_users_list(data):
    rows = []
    for u in store.all_users(data)[:20]:
        icon  = "✅" if u["allowed"] else ("🚫" if u["blocked"] else "⏳")
        rows.append([InlineKeyboardButton(f"{icon} {u['name'][:20]}", callback_data=f"adm:user:{u['id']}")])
    rows.append([InlineKeyboardButton("⏳ Pending Only", callback_data="adm:pending")])
    rows.append([InlineKeyboardButton("⬅️ Back",         callback_data="adm:main")])
    return InlineKeyboardMarkup(rows)

def kb_user_actions(uid, u):
    rows = []
    if not u["allowed"] and not u["blocked"]:
        rows.append([
            InlineKeyboardButton("✅ Approve", callback_data=f"adm:approve:{uid}"),
            InlineKeyboardButton("❌ Deny",    callback_data=f"adm:deny:{uid}"),
        ])
    elif u["allowed"] and not u["blocked"]:
        rows.append([
            InlineKeyboardButton("🚫 Block",  callback_data=f"adm:block:{uid}"),
            InlineKeyboardButton("❌ Revoke", callback_data=f"adm:revoke:{uid}"),
        ])
    elif u["blocked"]:
        rows.append([InlineKeyboardButton("✅ Unblock", callback_data=f"adm:unblock:{uid}")])
    rows.append([InlineKeyboardButton("⬅️ Back", callback_data="adm:users")])
    return InlineKeyboardMarkup(rows)


# ── /admin command ────────────────────────────────────────────────────────────

@admin_only
async def cmd_admin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    data = store.load()
    s    = data["settings"]
    lp   = "✅ ON" if s.get("llamaparse_on") else "🔴 OFF"
    keys = engine.status()
    avail = sum(1 for k in keys if k["available"])
    text = (
        "👑 *Admin Panel*\n\n"
        f"🤖 Bot: {'✅ ON' if s['bot_enabled'] else '🔴 OFF'}\n"
        f"📄 LlamaParse: {lp}\n"
        f"👥 Users: {len(data['users'])} | ⏳ Pending: {len(store.pending_users(data))}\n"
        f"🔑 API Keys: {avail}/{len(keys)} available\n"
        f"📊 Today: {data['stats']['by_date'].get(datetime.now().strftime('%Y-%m-%d'), 0)} requests\n"
    )
    await update.message.reply_text(text, parse_mode="Markdown", reply_markup=kb_main())


# ── Callback router ────────────────────────────────────────────────────────────

@admin_only
async def handle_admin_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q    = update.callback_query
    cb   = q.data
    await q.answer()
    data  = store.load()
    parts = cb.split(":")

    # ── Main ──────────────────────────────────────────────────────────────────
    if cb == "adm:main":
        s    = data["settings"]
        lp   = "✅ ON" if s.get("llamaparse_on") else "🔴 OFF"
        keys = engine.status()
        avail = sum(1 for k in keys if k["available"])
        text = (
            "👑 *Admin Panel*\n\n"
            f"🤖 Bot: {'✅ ON' if s['bot_enabled'] else '🔴 OFF'}\n"
            f"📄 LlamaParse: {lp}\n"
            f"👥 Users: {len(data['users'])} | ⏳ Pending: {len(store.pending_users(data))}\n"
            f"🔑 API Keys: {avail}/{len(keys)} available\n"
            f"📊 Today: {data['stats']['by_date'].get(datetime.now().strftime('%Y-%m-%d'), 0)} requests\n"
        )
        await q.edit_message_text(text, parse_mode="Markdown", reply_markup=kb_main())

    # ── Users ─────────────────────────────────────────────────────────────────
    elif cb == "adm:users":
        users   = store.all_users(data)
        pending = store.pending_users(data)
        await q.edit_message_text(
            f"👥 *Users* — {len(users)} total\n"
            f"⏳ Pending: {len(pending)} | ✅ Approved: {sum(1 for u in users if u['allowed'])}\n"
            f"🚫 Blocked: {sum(1 for u in users if u['blocked'])}",
            parse_mode="Markdown", reply_markup=kb_users_list(data))

    elif cb == "adm:pending":
        pending = store.pending_users(data)
        if not pending:
            await q.edit_message_text("✅ No pending users.", reply_markup=kb_back("users"))
            return
        rows = [[InlineKeyboardButton(f"⏳ {u['name'][:20]}", callback_data=f"adm:user:{u['id']}")] for u in pending[:15]]
        rows.append([InlineKeyboardButton("⬅️ Back", callback_data="adm:users")])
        await q.edit_message_text(f"⏳ *{len(pending)} Pending*", parse_mode="Markdown",
                                  reply_markup=InlineKeyboardMarkup(rows))

    elif parts[1] == "user" and len(parts) == 3:
        uid_int = int(parts[2])
        u = store.get_user(data, uid_int)
        if not u:
            await q.edit_message_text("User not found.", reply_markup=kb_back("users"))
            return
        status = "✅ Approved" if u["allowed"] else ("🚫 Blocked" if u["blocked"] else "⏳ Pending")
        await q.edit_message_text(
            f"👤 *{u['name']}*\n@{u['username'] or 'N/A'} | ID: `{uid_int}`\n"
            f"Status: {status} | Requests: {u['requests']}\nJoined: {u['joined'][:10]}",
            parse_mode="Markdown", reply_markup=kb_user_actions(uid_int, u))

    elif parts[1] == "approve":
        uid_int = int(parts[2])
        store.set_allowed(data, uid_int, True)
        u = store.get_user(store.load(), uid_int)
        await q.edit_message_text(f"✅ *{u['name']}* approved!", parse_mode="Markdown", reply_markup=kb_back("users"))
        try:
            welcome = data["settings"]["welcome_msg"]
            await context.bot.send_message(uid_int, f"🎉 Your access has been *approved*!\n\n{welcome}", parse_mode="Markdown")
        except Exception: pass

    elif parts[1] == "deny":
        uid_int = int(parts[2])
        store.set_blocked(data, uid_int, True)
        await q.edit_message_text("❌ User denied.", reply_markup=kb_back("users"))

    elif parts[1] == "block":
        uid_int = int(parts[2])
        store.set_blocked(data, uid_int, True)
        await q.edit_message_text("🚫 User blocked.", reply_markup=kb_back("users"))

    elif parts[1] == "unblock":
        uid_int = int(parts[2])
        store.set_blocked(data, uid_int, False)
        await q.edit_message_text("✅ User unblocked.", reply_markup=kb_back("users"))

    elif parts[1] == "revoke":
        uid_int = int(parts[2])
        store.set_allowed(data, uid_int, False)
        await q.edit_message_text("❌ Access revoked.", reply_markup=kb_back("users"))

    # ── Settings ──────────────────────────────────────────────────────────────
    elif cb == "adm:settings":
        await q.edit_message_text("⚙️ *Settings*", parse_mode="Markdown", reply_markup=kb_settings(data))

    elif cb == "adm:toggle_bot":
        data["settings"]["bot_enabled"] = not data["settings"]["bot_enabled"]
        store.save(data)
        await q.edit_message_text(f"🤖 Bot: {'✅ ON' if data['settings']['bot_enabled'] else '🔴 OFF'}",
                                  parse_mode="Markdown", reply_markup=kb_settings(store.load()))

    elif cb == "adm:toggle_whitelist":
        data["settings"]["whitelist_only"] = not data["settings"]["whitelist_only"]
        store.save(data)
        await q.edit_message_text(f"🔒 Whitelist: {'ON ✅' if data['settings']['whitelist_only'] else 'OFF 🔴'}",
                                  parse_mode="Markdown", reply_markup=kb_settings(store.load()))

    elif parts[1] == "set" and len(parts) == 3:
        field = parts[2]
        prompts = {
            "max_file_mb":   "📦 Enter max file size in MB (e.g. 15):",
            "cooldown_sec":  "⏱️ Enter cooldown seconds (e.g. 30):",
            "daily_limit":   "📅 Enter daily request limit (e.g. 20):",
            "gemini_model":  "🤖 Enter model name:\n• `gemini-2.0-flash` (fast, free)\n• `gemini-1.5-pro` (smarter)\n• `gemini-2.0-flash-thinking-exp` (reasoning)",
            "welcome_msg":   "✏️ Enter new welcome message:",
        }
        context.user_data["awaiting_setting"] = field
        await q.edit_message_text(prompts.get(field, f"Enter value for {field}:") + "\n\n_(Send /cancel to cancel)_",
                                  parse_mode="Markdown", reply_markup=kb_back("settings"))

    # ── AI Mode ───────────────────────────────────────────────────────────────
    elif cb == "adm:aimode":
        await q.edit_message_text("🤖 *AI Mode Settings*\nControl how documents are processed.",
                                  parse_mode="Markdown", reply_markup=kb_aimode(data))

    elif cb == "adm:toggle_llama":
        if not LLAMA_API_KEY:
            await q.answer("⚠️ LLAMA_API_KEY not set in .env!", show_alert=True)
            return
        data["settings"]["llamaparse_on"] = not data["settings"].get("llamaparse_on", False)
        store.save(data)
        state = "ON ✅" if data["settings"]["llamaparse_on"] else "OFF 🔴"
        await q.edit_message_text(f"📄 LlamaParse: *{state}*\n\n{'✅ PDFs will now use LlamaParse first for better extraction.' if data['settings']['llamaparse_on'] else '🔄 Reverted to Gemini-only extraction.'}",
                                  parse_mode="Markdown", reply_markup=kb_aimode(store.load()))

    elif cb == "adm:llama_info":
        await q.edit_message_text(
            "📄 *What is LlamaParse?*\n\n"
            "LlamaParse is a specialized PDF parsing service by LlamaIndex.\n\n"
            "*When ON (PDFs only):*\n"
            "• Much better table extraction\n"
            "• Preserves complex layouts\n"
            "• Better multi-column handling\n"
            "• Free: 1,000 pages/day\n\n"
            "*When OFF:*\n"
            "• Uses Gemini vision only\n"
            "• Works for both images and PDFs\n"
            "• No extra API key needed\n\n"
            "📌 Get key free at: cloud.llamaindex.ai\n"
            "Add `LLAMA_API_KEY=...` to your `.env`",
            parse_mode="Markdown", reply_markup=kb_back("aimode"))

    # ── Stats ─────────────────────────────────────────────────────────────────
    elif cb == "adm:stats":
        s     = data["stats"]
        today = datetime.now().strftime("%Y-%m-%d")
        top5  = sorted(s["by_user"].items(), key=lambda x: x[1], reverse=True)[:5]
        top_str = ""
        for uid_str, cnt in top5:
            u = data["users"].get(uid_str, {})
            top_str += f"  • {u.get('name', uid_str)}: {cnt}\n"
        lp_used = s.get("llamaparse_used", 0)
        await q.edit_message_text(
            f"📊 *Statistics*\n\n"
            f"Total: *{s['total']}* | ✅ {s['success']} | ❌ {s['failed']}\n"
            f"Today: *{s['by_date'].get(today, 0)}*\n"
            f"📄 LlamaParse used: *{lp_used}* times\n\n"
            f"🏆 *Top Users:*\n{top_str or '  None yet'}",
            parse_mode="Markdown", reply_markup=kb_back())

    # ── API Keys ──────────────────────────────────────────────────────────────
    elif cb == "adm:keys":
        lines = []
        for ks in engine.status():
            avail  = "✅ Available" if ks["available"] else "🔴 Exhausted"
            retry  = f"\n  ↳ retry: {ks['retry_after'][:16]}" if ks.get("retry_after") else ""
            lines.append(f"Key #{ks['index']}: {avail} | {ks['used']} uses{retry}")
        llama_status = f"\n📄 LlamaParse key: {'✅ Set' if LLAMA_API_KEY else '❌ Not set'}"
        await q.edit_message_text(
            "🔑 *API Keys*\n\n" + "\n".join(lines) + llama_status,
            parse_mode="Markdown", reply_markup=kb_back())

    # ── Broadcast ─────────────────────────────────────────────────────────────
    elif cb == "adm:broadcast":
        await q.edit_message_text("📢 *Broadcast*", parse_mode="Markdown",
                                  reply_markup=InlineKeyboardMarkup([
                                      [InlineKeyboardButton("📢 All Users",  callback_data="adm:bcast_all")],
                                      [InlineKeyboardButton("📨 One User",   callback_data="adm:bcast_one")],
                                      [InlineKeyboardButton("⬅️ Back",       callback_data="adm:main")],
                                  ]))

    elif cb == "adm:bcast_all":
        context.user_data["awaiting_broadcast"] = "all"
        await q.edit_message_text("📢 Type broadcast message:\n_(Send /cancel to cancel)_",
                                  parse_mode="Markdown", reply_markup=kb_back("broadcast"))

    elif cb == "adm:bcast_one":
        context.user_data["awaiting_broadcast"] = "one_id"
        await q.edit_message_text("📨 Enter user Telegram ID:\n_(Send /cancel to cancel)_",
                                  parse_mode="Markdown", reply_markup=kb_back("broadcast"))

    # ── Security ──────────────────────────────────────────────────────────────
    elif cb == "adm:security":
        blocked = [u for u in store.all_users(data) if u.get("blocked")]
        rows = [[InlineKeyboardButton(f"🚫 {u['name'][:20]}", callback_data=f"adm:user:{u['id']}")] for u in blocked[:15]]
        rows.append([InlineKeyboardButton("⬅️ Back", callback_data="adm:main")])
        await q.edit_message_text(f"🔒 *Security*\n🚫 Blocked: {len(blocked)} users",
                                  parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(rows))

    # ── Data ──────────────────────────────────────────────────────────────────
    elif cb == "adm:data":
        await q.edit_message_text("🗂️ *Data Management*", parse_mode="Markdown",
                                  reply_markup=InlineKeyboardMarkup([
                                      [InlineKeyboardButton("📤 Export Users",  callback_data="adm:export_users")],
                                      [InlineKeyboardButton("🔄 Reset Stats",   callback_data="adm:reset_stats")],
                                      [InlineKeyboardButton("⬅️ Back",          callback_data="adm:main")],
                                  ]))

    elif cb == "adm:export_users":
        buf = io.BytesIO(json.dumps(data["users"], indent=2).encode())
        await context.bot.send_document(update.effective_chat.id, document=buf,
                                        filename="users_export.json", caption="📤 Users export")
        await q.edit_message_text("✅ Exported!", reply_markup=kb_back("data"))

    elif cb == "adm:reset_stats":
        data["stats"] = {"total": 0, "success": 0, "failed": 0, "by_date": {}, "by_user": {}, "llamaparse_used": 0}
        store.save(data)
        await q.edit_message_text("✅ Stats reset.", reply_markup=kb_back("data"))


# ── Text input handler ────────────────────────────────────────────────────────

@admin_only
async def handle_admin_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> bool:
    text = update.message.text.strip()

    if text == "/cancel":
        context.user_data.pop("awaiting_setting",  None)
        context.user_data.pop("awaiting_broadcast", None)
        context.user_data.pop("broadcast_target",  None)
        await update.message.reply_text("❌ Cancelled.")
        return True

    if "awaiting_setting" in context.user_data:
        field = context.user_data.pop("awaiting_setting")
        data  = store.load()
        if field in ("max_file_mb", "cooldown_sec", "daily_limit"):
            try:
                data["settings"][field] = int(text)
            except ValueError:
                await update.message.reply_text("⚠️ Please enter a number.")
                context.user_data["awaiting_setting"] = field
                return True
        else:
            data["settings"][field] = text
        store.save(data)
        await update.message.reply_text(f"✅ *{field}* → `{text}`", parse_mode="Markdown")
        return True

    if "awaiting_broadcast" in context.user_data:
        stage = context.user_data["awaiting_broadcast"]
        if stage == "one_id":
            try:
                context.user_data["broadcast_target"]  = int(text)
                context.user_data["awaiting_broadcast"] = "one_msg"
                await update.message.reply_text("📨 Now type the message:")
            except ValueError:
                await update.message.reply_text("⚠️ Invalid ID.")
            return True
        if stage in ("all", "one_msg"):
            data = store.load()
            targets = [int(uid) for uid, u in data["users"].items() if u.get("allowed")] \
                      if stage == "all" else [context.user_data.pop("broadcast_target")]
            context.user_data.pop("awaiting_broadcast", None)
            sent = failed = 0
            for tid in targets:
                try:
                    await context.bot.send_message(tid, text, parse_mode="Markdown")
                    sent += 1
                except Exception:
                    failed += 1
            await update.message.reply_text(f"📢 Done! ✅ {sent} sent | ❌ {failed} failed")
            return True

    return False
