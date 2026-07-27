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

# DIQQAT (2026-07-08 qo'shildi): Windows serverda (pm2/pythonw.exe orqali
# ishga tushirilganda) Python stdout/stderr kodировкasi standart ravishda
# tizim konsolining kodировkasiga (masalan kirillcha Windows'da cp1251)
# tushib qoladi — bu UTF-8 emoji (✅, ⚠️ va h.k.) bor har qanday print()
# chaqiruvida "UnicodeEncodeError: 'charmap' codec can't encode character"
# xatosiga olib keladi (masalan Generate_Asosiy_order.py'dagi
# print(f"✅ Real ma'lumot...") shu sababli butun "Buyurtma Excel olish"
# jarayonini qulatib qo'ygan edi). Kodda emoji bilan print() qiluvchi ko'p
# joy bor (main.py, Generate_Asosiy_order.py va h.k.), shuning uchun buni
# BIR MARTA, dasturning eng boshida, butun jarayon uchun tuzatamiz —
# har bir print() joyini alohida qidirib tuzatishga hojat qolmaydi.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from telegram.ext import (
    Application, CommandHandler, CallbackQueryHandler,
    MessageHandler, filters, ContextTypes,
)
from telegram import Update
from telegram.request import HTTPXRequest

from datetime import time as _time

from config import BOT_TOKEN, SUPER_ADMIN_ID, logger, xlsx_refresh
from handlers import (
    start, callback_handler, text_keldi, fayl_keldi, adduser_cmd,
    removeuser_cmd, users_cmd, chatid_cmd, perexod_kunlik_tekshiruv,
    addadmin_cmd, removeadmin_cmd, bot_holat_zaxira_yubor,
)


# 2026-07-25 (Xavfsizlik/barqarorlik auditi, Huzayfa so'rovi bilan):
# ILGARI HECH QANDAY global xato-ushlagich (error handler) YO'Q edi —
# har bir handler o'zining try/except'iga tayanardi, lekin biror joyda
# kutilmagan (bashorat qilinmagan) xatolik chiqsa, python-telegram-bot
# uni ICHKI ravishda log qilib, botni "o'zi" davom ettiradi (Application
# butun jarayonni qulatib qo'ymaydi) — LEKIN foydalanuvchiga HECH QANDAY
# javob ketmaydi (jimgina "osilib qoladi") va ADMIN xato haqida umuman
# BILMAYDI, faqat serverdagi konsol logini qo'lda ochib ko'rsagina
# ko'radi. Bu funksiya ikkalasini ham tuzatadi: (1) foydalanuvchiga
# tushunarli xabar, (2) super adminga darhol xato haqida qisqa xabar
# (Telegram xabar hajmi chegarasi uchun 1500 belgigacha kesilgan holda).
async def global_xato_ushlagich(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    logger.error("Kutilmagan xatolik:", exc_info=context.error)
    try:
        if isinstance(update, Update) and update.effective_message:
            await update.effective_message.reply_text(
                "❌ Kutilmagan xatolik yuz berdi. Admin xabardor qilindi — "
                "qayta urinib ko'ring yoki biroz kuting."
            )
    except Exception:
        pass
    if SUPER_ADMIN_ID:
        try:
            import traceback
            tb = "".join(traceback.format_exception(
                type(context.error), context.error, context.error.__traceback__
            ))
            uid_info = ""
            if isinstance(update, Update) and update.effective_user:
                uid_info = f"👤 uid={update.effective_user.id}\n"
            await context.bot.send_message(
                chat_id=SUPER_ADMIN_ID,
                text=f"🔴 Botda xatolik:\n{uid_info}```\n{tb[-1500:]}\n```",
                parse_mode="Markdown",
            )
        except Exception:
            logger.exception("Admin ga xato xabarini yuborib bo'lmadi")

try:
    from zoneinfo import ZoneInfo
    _TZ = ZoneInfo("Asia/Tashkent")
except Exception:
    _TZ = None


def main() -> None:
    xlsx_refresh(force=True)
    # 2026-07-27 (Huzayfa: adminga real "telegram.error.TimedOut" xatoligi
    # kelgan edi, `_send_message` ichida): python-telegram-bot standart
    # HTTPXRequest sozlamalari JUDA QATTIQ (read/write/connect_timeout —
    # bor-yo'g'i 5 soniya, pool_timeout — atigi 1 soniya). Bot Excel
    # fayllar (Buyurtma/Xitoy ostatka/Userlar ro'yxati) yuboradi — bular
    # ayniqsa sekinroq tarmoqda yoki bir vaqtning o'zida bir nechta
    # so'rov navbatga tursa (pool_timeout=1s) osongina shu qattiq
    # chegaradan chiqib ketadi va TimedOut bilan yiqiladi (garchi endi
    # global_xato_ushlagich buni ushlab, botni qulatmasa ham — foydalanuvchi
    # baribir "xatolik" xabarini ko'radi, holbuki bu shunchaki vaqtinchalik
    # tarmoq sekinligi edi). Yumshoqroq (lekin haddan tashqari uzun ham
    # emas — osilib qolishning oldini olish uchun) qiymatlarga o'zgartirildi.
    request = HTTPXRequest(
        connect_timeout=15.0,
        read_timeout=30.0,
        write_timeout=30.0,
        pool_timeout=10.0,
    )
    app = Application.builder().token(BOT_TOKEN).request(request).build()
    app.add_error_handler(global_xato_ushlagich)

    # 2026-07-16: "rasm yuborilgan, hali KELDI qilinmagan" konteynerlarni
    # har kuni tekshiradi — eslatma yuboradi, 4 kun o'tsa avtomatik KELDI
    # qiladi (handlers.py::perexod_kunlik_tekshiruv). Soat 09:00 (Toshkent).
    if app.job_queue is not None:
        app.job_queue.run_daily(
            perexod_kunlik_tekshiruv,
            time=_time(hour=9, minute=0, tzinfo=_TZ) if _TZ else _time(hour=4, minute=0),
        )
        # 2026-07-27 (Huzayfa: bot_holat/ zaxira nusxasiz edi, faqat
        # bitta admin bor — server buzilsa hammasi qaytarib bo'lmas
        # holda yo'qolardi): har kuni soat 00:10 (Toshkent, tinch vaqt)
        # bot_holat/ni zip qilib SUPER_ADMIN_ID'ga hujjat sifatida
        # yuboradi (handlers.py::bot_holat_zaxira_yubor).
        app.job_queue.run_daily(
            bot_holat_zaxira_yubor,
            time=_time(hour=0, minute=10, tzinfo=_TZ) if _TZ else _time(hour=19, minute=10),
        )
    else:
        logger.warning(
            "JobQueue mavjud emas — perexod eslatmasi ishlamaydi. "
            "requirements.txt'da 'python-telegram-bot[job-queue]' borligini tekshiring."
        )
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
    app.add_handler(CommandHandler("addadmin",   addadmin_cmd,   filters=private))
    app.add_handler(CommandHandler("removeadmin", removeadmin_cmd, filters=private))
    app.add_handler(CommandHandler("chatid",     chatid_cmd))
    app.add_handler(CallbackQueryHandler(callback_handler))
    app.add_handler(MessageHandler(private & filters.Document.ALL, fayl_keldi))
    app.add_handler(MessageHandler(private & filters.TEXT & ~filters.COMMAND, text_keldi))
    logger.info("Bot started.")
    app.run_polling()


if __name__ == "__main__":
    main()
