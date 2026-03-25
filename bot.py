"""
bot.py — Main entry point

Full pipeline:
  File received
    ↓
  [LlamaParse if PDF + enabled]  OR  [Gemini extraction]
    ↓
  Gemini style analysis (parallel to extraction)
    ↓
  Build StyleMap from style JSON + doc type
    ↓
  DOCX builder (sentence merge → block parse → render with styles)
    ↓
  Send .docx back
"""

import io, re, time, logging, asyncio
from datetime import datetime

from telegram import Update
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, filters, ContextTypes,
)
from telegram.error import NetworkError, TimedOut

import store, security
import admin as adm
from config import (
    TELEGRAM_TOKEN, ADMIN_ID, WEBHOOK_URL, WEBHOOK_PORT,
    WEBHOOK_PATH, USE_WEBHOOK, LOG_FILE,
)
from gemini import engine, extract_with_llamaparse
from detector import detect
from style_engine import build_style_map
from docx_builder import build_docx

# ─── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    handlers=[logging.FileHandler(LOG_FILE), logging.StreamHandler()],
    format="%(asctime)s | %(levelname)s | %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)
BOT_START = time.time()


# ══════════════════════════════════════════════════════════════════════════════
#  USER COMMANDS
# ══════════════════════════════════════════════════════════════════════════════

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    data = store.load()
    is_new = store.register_user(data, user.id, user.full_name, user.username or "")

    if security.is_admin(user.id):
        await update.message.reply_text(
            f"👑 *Admin Panel*\nUptime: `{_uptime()}`\n\nUse /admin to manage.",
            parse_mode="Markdown")
        return

    if not store.is_allowed(data, user.id, ADMIN_ID):
        await update.message.reply_text(
            "👋 *Welcome!*\n\nYour account is *pending approval*.\n"
            "An admin will review your request shortly.",
            parse_mode="Markdown")
        if is_new:
            try:
                await context.bot.send_message(
                    ADMIN_ID,
                    f"🔔 *New Access Request*\n\n👤 {user.full_name}\n"
                    f"🔗 @{user.username or 'N/A'}\n🆔 `{user.id}`\n\nUse /admin → Users.",
                    parse_mode="Markdown")
            except Exception: pass
        return

    welcome = data["settings"].get("welcome_msg", "👋 Send me a PDF or image!")
    await update.message.reply_text(welcome, parse_mode="Markdown")


async def cmd_status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    data = store.load()
    if not store.is_allowed(data, user.id, ADMIN_ID):
        return
    today_cnt = store.today_count(data, user.id)
    limit     = data["settings"]["daily_limit"]
    lp_on     = data["settings"].get("llamaparse_on", False)
    await update.message.reply_text(
        f"📊 *Status*\n\nRequests today: *{today_cnt}/{limit}*\n"
        f"AI mode: {'LlamaParse + Gemini' if lp_on else 'Gemini only'}\n"
        f"Uptime: `{_uptime()}`",
        parse_mode="Markdown")


# ══════════════════════════════════════════════════════════════════════════════
#  FILE HANDLERS
# ══════════════════════════════════════════════════════════════════════════════

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    store.register_user(store.load(), update.effective_user.id,
                        update.effective_user.full_name, update.effective_user.username or "")
    doc  = update.message.document
    mime = doc.mime_type or ""
    if "pdf" in mime:
        await _process_file(update, context, doc, "application/pdf", doc.file_name or "document.pdf")
    elif "image" in mime:
        await _process_file(update, context, doc, mime, doc.file_name or "image.jpg")
    else:
        await update.message.reply_text("⚠️ Please send a *PDF* or *image* file.", parse_mode="Markdown")


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    store.register_user(store.load(), update.effective_user.id,
                        update.effective_user.full_name, update.effective_user.username or "")
    photo = update.message.photo[-1]
    await _process_file(update, context, photo, "image/jpeg", "photo.jpg")


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if security.is_admin(update.effective_user.id):
        consumed = await adm.handle_admin_text(update, context)
        if consumed: return
    await update.message.reply_text(
        "Please send a *PDF* or *image* file.\nType /start for instructions.",
        parse_mode="Markdown")


# ══════════════════════════════════════════════════════════════════════════════
#  CORE PROCESSING PIPELINE
# ══════════════════════════════════════════════════════════════════════════════

async def _process_file(update, context, file_obj, mime_type: str, filename: str):
    user = update.effective_user
    data = store.load()

    # ── Auth & limits ─────────────────────────────────────────────────────────
    if not data["settings"]["bot_enabled"] and not security.is_admin(user.id):
        await update.message.reply_text("🔴 Bot is temporarily offline.")
        return

    if not store.is_allowed(data, user.id, ADMIN_ID):
        count = security.track_unauth(user.id)
        if count >= 5:
            store.set_blocked(data, user.id, True)
        await update.message.reply_text("⛔ You are not authorized. Use /start to request access.")
        return

    security.reset_unauth(user.id)

    err = security.validate_file(file_obj, data["settings"]["max_file_mb"])
    if err:
        await update.message.reply_text(err, parse_mode="Markdown")
        return

    wait = security.check_cooldown(user.id, data["settings"]["cooldown_sec"])
    if wait > 0:
        await update.message.reply_text(f"⏳ Please wait *{wait}s* before next file.", parse_mode="Markdown")
        return

    today_cnt = store.today_count(data, user.id)
    if today_cnt >= data["settings"]["daily_limit"] and not security.is_admin(user.id):
        await update.message.reply_text(
            f"📅 Daily limit of *{data['settings']['daily_limit']}* reached. Resets at midnight.",
            parse_mode="Markdown")
        return

    # ── Download ──────────────────────────────────────────────────────────────
    security.set_cooldown(user.id)
    store.increment_today(data, user.id)
    status_msg = await update.message.reply_text("⏳ Downloading your file…")

    try:
        file_tg    = await file_obj.get_file()
        buf        = io.BytesIO()
        await file_tg.download_to_memory(buf)
        file_bytes = buf.getvalue()

        # ── Step 1: Detect document type ─────────────────────────────────────
        await status_msg.edit_text("🔍 Analyzing document type…")
        available_keys = [ks for ks in engine.keys if ks.is_available()]
        first_key = available_keys[0].key if available_keys else ""
        doc_info = detect(file_bytes, mime_type, first_key, data["settings"]["gemini_model"])
        logger.info("Detected: %s | lang=%s | tables=%s | formulas=%s",
                    doc_info["type"], doc_info["lang"],
                    doc_info["has_tables"], doc_info["has_formulas"])

        # ── Step 2: Extract content ───────────────────────────────────────────
        is_pdf       = "pdf" in mime_type
        llamaparse_on = data["settings"].get("llamaparse_on", False)
        used_llama   = False
        extracted_text = None

        if is_pdf and llamaparse_on:
            await status_msg.edit_text("📄 Extracting with LlamaParse…")
            extracted_text = extract_with_llamaparse(file_bytes, filename)
            if extracted_text:
                used_llama = True
                logger.info("LlamaParse success: %d chars", len(extracted_text))
            else:
                logger.info("LlamaParse failed — falling back to Gemini")

        if not extracted_text:
            mode_label = "🤖 Extracting content with Gemini AI…"
            if used_llama is False and is_pdf and llamaparse_on:
                mode_label = "🤖 LlamaParse failed — using Gemini AI…"
            await status_msg.edit_text(mode_label)
            engine.set_model(data["settings"]["gemini_model"])
            extracted_text = engine.extract(file_bytes, mime_type, doc_info)

        logger.info("Extracted %d chars from '%s' (llama=%s)", len(extracted_text), filename, used_llama)

        # ── Step 3: Analyze visual styling ────────────────────────────────────
        await status_msg.edit_text("🎨 Analyzing document styling…")
        style_json = engine.analyze_style(file_bytes, mime_type)
        style_map  = build_style_map(doc_info["type"], style_json)
        logger.info("Style map built for type=%s, base_font=%s, base_size=%s",
                    doc_info["type"], style_map.base_font, style_map.base_size)

        # ── Step 4: Build DOCX ────────────────────────────────────────────────
        await status_msg.edit_text("📝 Building formatted DOCX…")
        docx_bytes  = build_docx(extracted_text, style_map)
        output_name = re.sub(r"\.[^.]+$", "", filename) + "_extracted.docx"

        # ── Step 5: Send ──────────────────────────────────────────────────────
        await status_msg.edit_text("📤 Sending your document…")
        method_note = " _(LlamaParse + Gemini)_" if used_llama else " _(Gemini AI)_"
        await update.message.reply_document(
            document=io.BytesIO(docx_bytes),
            filename=output_name,
            caption=f"✅ Done! `{filename}`{method_note}\nType: `{doc_info['type']}`",
            parse_mode="Markdown",
        )
        await status_msg.delete()
        store.record_request(store.load(), user.id, success=True, used_llamaparse=used_llama)

    except Exception as e:
        logger.exception("Processing failed for user %d", user.id)
        store.record_request(store.load(), user.id, success=False)
        await status_msg.edit_text(
            f"❌ Error:\n`{security.safe_error(e)}`\n\nPlease try again.",
            parse_mode="Markdown")


# ─── Error handler ────────────────────────────────────────────────────────────

async def error_handler(update, context: ContextTypes.DEFAULT_TYPE):
    if isinstance(context.error, (NetworkError, TimedOut)):
        logger.warning("Network hiccup: %s", type(context.error).__name__)
        return
    logger.exception("Unexpected error: %s", security.safe_error(context.error))


def _uptime():
    from datetime import timedelta
    d = timedelta(seconds=int(time.time() - BOT_START))
    h, r = divmod(int(d.total_seconds()), 3600)
    m, s = divmod(r, 60)
    return f"{h}h {m}m {s}s"


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    app = (
        Application.builder()
        .token(TELEGRAM_TOKEN)
        .read_timeout(60)
        .write_timeout(60)
        .connect_timeout(60)
        .pool_timeout(60)
        .build()
    )

    app.add_handler(CommandHandler("start",  cmd_start))
    app.add_handler(CommandHandler("status", cmd_status))
    app.add_handler(CommandHandler("admin",  adm.cmd_admin))
    app.add_handler(CallbackQueryHandler(adm.handle_admin_callback, pattern=r"^adm:"))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.PHOTO,        handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    app.add_error_handler(error_handler)

    if USE_WEBHOOK:
        logger.info("🌐 Webhook mode — %s", WEBHOOK_URL)
        app.run_webhook(
            listen="0.0.0.0",
            port=WEBHOOK_PORT,
            url_path=WEBHOOK_PATH,
            webhook_url=f"{WEBHOOK_URL}{WEBHOOK_PATH}",
        )
    else:
        logger.info("🔄 Polling mode")
        app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
