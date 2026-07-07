"""
Bot.py — NEJAVIYKA Telegram Bot
================================
Entry point. Barcha logika alohida modullarda:
  config.py    — sozlamalar, paths, cache
  texts.py     — tarjimalar
  keyboards.py — klaviaturalar
  parsers.py   — xitoy Excel parse
  services.py  — buyurtma, inventar, konteyner
  ui.py        — ekranlar, grafik, animatsiya
  handlers.py  — Telegram handlerlari
"""
import sys
from pathlib import Path

# NEJAVIYKA papkasi sys.path da bo'lsin
_DIR = Path(__file__).resolve().parent
if str(_DIR) not in sys.path:
    sys.path.insert(0, str(_DIR))

from telegram.ext import (
    Application, CommandHandler, CallbackQueryHandler,
    MessageHandler, filters,
)

from config import BOT_TOKEN, logger, xlsx_refresh
from handlers import start, callback_handler, text_keldi, fayl_keldi, adduser_cmd, removeuser_cmd, users_cmd, chatid_cmd


def main() -> None:
    xlsx_refresh(force=True)
    app = Application.builder().token(BOT_TOKEN).build()
    # DIQQAT: botning butun menyu/navigatsiya tizimi FAQAT shaxsiy (private)
    # chat uchun mo'ljallangan — "Kelgan yuklar" kabi guruhlar faqat botdan
    # xabar QABUL qiladi, hech qachon botga buyruq/matn YUBORMAYDI. Shu
    # sababli bu handlerlar filters.ChatType.PRIVATE bilan cheklandi — aks
    # holda guruhdagi har qanday odam yozgan har qanday xabarga (yoki hujjat
    # yuborsa) bot javob berib, guruhni chalg'itib yuborardi.
    # /chatid bundan MUSTASNO — uning yagona vazifasi aynan guruh/topic
    # ID'sini olish, shuning uchun u ATAYLAB har qanday chatda ishlaydi
    # (lekin ADMIN_IDS bilan ichkarida cheklangan — handlers.py'ga qarang).
    private = filters.ChatType.PRIVATE
    app.add_handler(CommandHandler("start",      start,       filters=private))
    app.add_handler(CommandHandler("adduser",    adduser_cmd, filters=private))
    app.add_handler(CommandHandler("removeuser", removeuser_cmd, filters=private))
    app.add_handler(CommandHandler("users",      users_cmd,   filters=private))
    app.add_handler(CommandHandler("chatid",     chatid_cmd))
    app.add_handler(CallbackQueryHandler(callback_handler))
    app.add_handler(MessageHandler(private & filters.Document.ALL, fayl_keldi))
    app.add_handler(MessageHandler(private & filters.TEXT & ~filters.COMMAND, text_keldi))
    logger.info("Bot started.")
    app.run_polling()


if __name__ == "__main__":
    main()
