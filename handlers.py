"""
handlers.py — Telegram handler funksiyalar
  - start           : /start komandasi
  - callback_handler: inline tugmalar
  - text_keldi      : reply keyboard tugmalari va matn
  - fayl_keldi      : .xlsx fayl qabul qilish
"""
import asyncio
import logging
import os
import sys
from io import BytesIO
from pathlib import Path

from telegram import Update

logger = logging.getLogger(__name__)
from telegram.ext import ContextTypes

import config
from config import BOT_HOLAT_DIR, CH_KEY, BACK_MAP, BASE_DIR, KONT_DIR, XITOY_PARSED_DIR, ADMIN_IDS, SUPER_ADMIN_ID, SUPPORT_PHONE, SUPPORT_USERNAME
from config import xlsx_refresh as _xlsx_refresh
from kont_rasm import generate_kelgan_rasm
from texts import t
from keyboards import (
    main_kb, order_kb, order_channel_kb, load_kb, load_channel_kb,
    settings_kb, search_kb, til_ikb, konteyner_kb,
    xitoy_sorash_ikb, xitoy_mavjud_ikb, xitoy_yana_ikb,
    tozala_kanal_ikb, tozala_tasdiq_ikb, zakaz_tasdiq_ikb,
    grafik_kat_ikb, kont_tasdiq_ikb, boglanish_ikb,
)
from telegram import InlineKeyboardMarkup, InlineKeyboardButton, InputMediaPhoto
from services import (
    buyurtma_yuklash, buyurtma_saqlash, buyurtma_tozala,
    pending_saqlash, pending_yuklash, pending_tozala,
    draft_saqlash, draft_yuklash,
    xitoy_yuklash, xitoy_saqlash, buyurtma_tekshir,
    zakaz_preview_text,
    grafik_qidirish, vazn_lookup_yangilash,
    qidiruv_olish, qidiruv_text,
    konteyner_tarix_olish, konteyner_tarix_qoshish, konteyner_tarix_kalit,
    kirish_ruxsati, whitelist_qosh, whitelist_ochir, whitelist_yuklash,
)
from parsers import xitoy_ostatka_oqi
from keyboards import grafik_tovar_ikb
from ui import (
    build_screen, go_screen, get_action,
    yuklash_animatsiya, grafik_ko_rsatish,
    draft_buyurtma_yubor, yolda_ko_rish,
)
from yuklatish_rejasi import main_with_data

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid  = update.message.from_user.id
    user = update.message.from_user
    context.user_data["user_id"] = uid

    # ── Kirish tekshiruvi ──────────────────────────────────────────────────────
    if not kirish_ruxsati(uid):
        # Adminlarga xabar yuborish
        full_name = (user.full_name or "").strip() or "Noma'lum"
        username  = f"@{user.username}" if user.username else "(username yo'q)"
        notif = (
            f"🔔 *Yangi kirish so'rovi*\n\n"
            f"👤 {full_name} {username}\n"
            f"🆔 `{uid}`\n\n"
            f"Qo'shish uchun: `/adduser {uid}`"
        )
        logger.info(f"Kirish so'rovi: uid={uid}, SUPER_ADMIN_ID={SUPER_ADMIN_ID}")
        if SUPER_ADMIN_ID:
            try:
                await context.bot.send_message(
                    chat_id=SUPER_ADMIN_ID, text=notif, parse_mode="Markdown"
                )
                logger.info(f"Admin ga xabar yuborildi: {SUPER_ADMIN_ID}")
            except Exception as e:
                logger.error(f"Admin ga xabar yuborish XATO: {e}")
        await update.message.reply_text(
            "⛔ Kechirasiz, sizda kirish huquqi yo'q.\n\n"
            "Admin bilan bog'laning — so'rovingiz yuborildi."
        )
        return

    lang = context.user_data.get("lang")
    if lang:
        await go_screen(update.message, context, "main")
    else:
        await update.message.reply_text(
            "🇺🇿 Тилни танланг / Tilni tanlang:",
            reply_markup=til_ikb(),
        )


async def adduser_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/adduser <id> — whitelist ga qo'shish (faqat super admin)."""
    uid = update.message.from_user.id
    if uid != SUPER_ADMIN_ID:
        await update.message.reply_text("⛔ Faqat adminlar uchun.")
        return
    args = context.args
    if not args or not args[0].isdigit():
        await update.message.reply_text("❓ Format: `/adduser 123456789`", parse_mode="Markdown")
        return
    new_id = int(args[0])
    added  = whitelist_qosh(new_id)
    if added:
        await update.message.reply_text(f"✅ `{new_id}` whitelist ga qo'shildi.", parse_mode="Markdown")
        try:
            await context.bot.send_message(
                chat_id=new_id,
                text="✅ Botga kirish huquqi berildi! /start bosing."
            )
        except Exception:
            pass
    else:
        await update.message.reply_text(f"ℹ️ `{new_id}` allaqachon ro'yxatda.", parse_mode="Markdown")


async def removeuser_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/removeuser <id> — whitelist dan o'chirish (faqat super admin)."""
    uid = update.message.from_user.id
    if uid != SUPER_ADMIN_ID:
        await update.message.reply_text("⛔ Faqat adminlar uchun.")
        return
    args = context.args
    if not args or not args[0].isdigit():
        await update.message.reply_text("❓ Format: `/removeuser 123456789`", parse_mode="Markdown")
        return
    rem_id = int(args[0])
    removed = whitelist_ochir(rem_id)
    if removed:
        await update.message.reply_text(f"✅ `{rem_id}` ro'yxatdan o'chirildi.", parse_mode="Markdown")
    else:
        await update.message.reply_text(f"ℹ️ `{rem_id}` ro'yxatda topilmadi.", parse_mode="Markdown")


async def users_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/users — whitelist va adminlar ro'yxati (faqat super admin)."""
    uid = update.message.from_user.id
    if uid != SUPER_ADMIN_ID:
        await update.message.reply_text("⛔ Faqat super admin uchun.")
        return
    wl = whitelist_yuklash()
    lines = ["👥 *Foydalanuvchilar ro'yxati:*\n"]
    lines.append("🔑 *Adminlar (ADMIN\\_IDS):*")
    for aid in sorted(ADMIN_IDS):
        marker = " ← siz" if aid == SUPER_ADMIN_ID else ""
        lines.append(f"  `{aid}`{marker}")
    lines.append("")
    lines.append("✅ *Whitelist:*")
    if wl:
        for wid in sorted(wl):
            lines.append(f"  `{wid}`")
    else:
        lines.append("  _(bo'sh)_")
    lines.append(f"\nJami whitelist: *{len(wl)}* ta")
    await update.message.reply_text("\n".join(lines), parse_mode="Markdown")


async def chatid_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/chatid — joriy chat (guruh/topic) ID sini ko'rsatadi.
    Botni istalgan guruhga (xususiy, oddiy yoki topic'li) qo'shib, shu yerda
    /chatid deb yozsangiz — .env uchun kerakli KELGAN_YUKLAR_CHAT_ID va
    (agar topic bo'lsa) KELGAN_YUKLAR_TOPIC_ID qiymatlarini darhol ko'rasiz.
    Telegram linkidan (t.me/c/...) qidirishga hojat qolmaydi — bu oddiy
    (topic'siz) guruhlarda umuman ishlamaydi."""
    uid = update.message.from_user.id
    if ADMIN_IDS and uid not in ADMIN_IDS:
        return  # guruhdagi boshqa a'zolarga javob bermaymiz
    chat = update.effective_chat
    msg  = update.message
    lines = [
        "🆔 *Chat ma'lumotlari:*",
        f"Chat ID: `{chat.id}`",
        f"Turi: {chat.type}",
    ]
    if chat.title:
        lines.append(f"Nomi: {chat.title}")
    thread_id = getattr(msg, "message_thread_id", None)
    if thread_id:
        lines.append(f"Topic (mavzu) ID: `{thread_id}`")
    else:
        lines.append("_Bu xabar biror topic/mavzuga tegishli emas — oddiy guruh bo'lsa shunday bo'lishi kerak._")
    await msg.reply_text("\n".join(lines), parse_mode="Markdown")


async def callback_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Til tanlash va Xitoy sorash uchun (inline keyboard)."""
    query = update.callback_query
    await query.answer()
    uid = query.from_user.id
    if not kirish_ruxsati(uid):
        await query.answer("⛔ Kirish huquqi yo'q.", show_alert=True)
        return

    if query.data.startswith("lang:"):
        lang = query.data.split(":")[1]
        context.user_data["lang"]   = lang
        context.user_data["screen"] = "main"
        await query.edit_message_text(
            "✅ Тил танланди!" if lang == "cyr" else "✅ Til tanlandi!"
        )
        await go_screen(query.message, context, "main")

    elif query.data == "lang_pick":
        await query.message.reply_text("🇺🇿 Тилни танланг:", reply_markup=til_ikb())

    elif query.data.startswith("xitoy:"):
        # xitoy:ha/yoq/ishlatsin/yangi:{kanal}
        parts    = query.data.split(":")
        decision = parts[1]
        kanal    = parts[2]
        lang     = context.user_data.get("lang", "cyr")

        if decision in ("yoq", "hisob_siz"):
            # Xitoy ostatka hisobga olinmaydi — bo'sh marker saqlaymiz
            # (keyingi kirishda "xitoy_sorash" emas "xitoy_mavjud" ko'rinsin)
            xitoy_saqlash(kanal, {}, {})
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            await draft_buyurtma_yubor(query.message, context, kanal, lang,
                                        xitoy_ostatka={})

        elif decision == "ha":
            # 2-fayl flow: avval Труба/Профиль so'rab, keyin Лист so'raymiz
            context.user_data.pop("xitoy_tp_data", None)
            context.user_data["kutilmoqda"] = ("xitoy_tp", kanal)
            await query.edit_message_text(
                t(lang, "xitoy_fayl_kut_tp"), parse_mode="Markdown"
            )

        elif decision == "ishlatsin":
            mavjud = xitoy_yuklash(kanal)
            xitoy_map = mavjud.get("tovarlar", {}) if mavjud else {}
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            await draft_buyurtma_yubor(query.message, context, kanal, lang,
                                        xitoy_ostatka=xitoy_map)

        elif decision == "yangi":
            # Yangi ostatka yuklashda: eski xitoy JSON + tasdiqlangan buyurtma ikkalasi o'chadi.
            p = BOT_HOLAT_DIR / f"xitoy_{kanal}.json"
            if p.exists():
                p.unlink()
            buyurtma_tozala(kanal)
            context.user_data.pop("xitoy_tp_data", None)
            # 2-fayl flow: avval TP so'ra
            context.user_data["kutilmoqda"] = ("xitoy_tp", kanal)
            await query.edit_message_text(
                t(lang, "xitoy_fayl_kut_tp"), parse_mode="Markdown"
            )

        elif decision == "yana_f":
            # Ko'p-fayl yig'ishda keyingi faylni kutish (konteyner/yuklatish flow)
            context.user_data["kutilmoqda"] = ("xitoy_fayl", kanal)
            await query.edit_message_text(
                t(lang, "xitoy_fayl_kut"), parse_mode="Markdown"
            )

        elif decision == "tayyor":
            # xitoy_ostatka_fayl (buyurtma) uchun → draft Excel
            context.user_data.pop("kutilmoqda", None)
            mavjud    = xitoy_yuklash(kanal)
            xitoy_map = mavjud.get("tovarlar", {}) if mavjud else {}
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            await draft_buyurtma_yubor(query.message, context, kanal, lang,
                                        xitoy_ostatka=xitoy_map)

        elif decision == "hisoblash":
            # xitoy_fayl (yuklatish rejasi) uchun → main_with_data
            context.user_data.pop("kutilmoqda", None)
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            ombor_akkum = context.user_data.get("ombor_akkum", {})
            vazn_akkum  = context.user_data.get("vazn_akkum", {})
            async with yuklash_animatsiya(
                query.message, context,
                text_cyr="Юклатиш режаси тайёрланмоқда",
                text_lat="Yuklatish rejasi tayyorlanmoqda",
            ):
                try:
                    import sys as _sys, importlib.util as _ilu
                    _nd = Path(__file__).resolve().parent
                    for _dep in ("vazn_hisobla", "Yuklama_optimal", "yuklatish_rejasi"):
                        if _dep not in _sys.modules:
                            _sp = _ilu.spec_from_file_location(_dep, _nd / f"{_dep}.py")
                            _m  = _ilu.module_from_spec(_sp); _sys.modules[_dep] = _m; _sp.loader.exec_module(_m)
                    xlsx_path = _sys.modules["yuklatish_rejasi"].main_with_data(kanal, ombor_akkum, xitoy_vazn=vazn_akkum)
                except Exception as e:
                    await query.message.reply_text(f"❌ Yuklatish rejasi xato:\n{str(e)[:300]}")
                    return
            context.user_data.pop("xitoy_akkum", None)
            context.user_data.pop("ombor_akkum", None)
            context.user_data.pop("vazn_akkum", None)
            if xlsx_path:
                with open(xlsx_path.split("|")[-1] if "|" in xlsx_path else xlsx_path, "rb") as f:
                    await query.message.reply_document(document=f, filename="Yuklatish_rejasi.xlsx")

    # ── Tozalash flow ─────────────────────────────────────────
    elif query.data.startswith("tozala_b:") or query.data.startswith("tozala_x:"):
        # Kanal tanlandi, tasdiqlash so'rash
        tip, kanal = query.data.split(":", 1)[0].split("_")[1], query.data.split(":")[1]
        lang = context.user_data.get("lang", "cyr")
        ch   = t(lang, CH_KEY.get(kanal, kanal))

        if tip == "b":
            data = buyurtma_yuklash(kanal)
            n    = len(data.get("buyurtmalar", [])) if data else 0
            sana = data.get("sana", "?") if data else "?"
            if not data or n == 0:
                await query.edit_message_text(
                    t(lang, "tozala_topilmadi").format(ch=ch),
                    parse_mode="Markdown"
                )
                return
            await query.edit_message_text(
                t(lang, "tozala_tasdiq_buy").format(ch=ch, n=n, sana=sana),
                parse_mode="Markdown",
                reply_markup=tozala_tasdiq_ikb(lang, "b", kanal),
            )
        else:
            data = xitoy_yuklash(kanal)
            n    = len(data.get("tovarlar", {})) if data else 0
            sana = data.get("sana", "?") if data else "?"
            if not data or n == 0:
                await query.edit_message_text(
                    t(lang, "tozala_topilmadi").format(ch=ch),
                    parse_mode="Markdown"
                )
                return
            await query.edit_message_text(
                t(lang, "tozala_tasdiq_xitoy").format(ch=ch, n=n, sana=sana),
                parse_mode="Markdown",
                reply_markup=tozala_tasdiq_ikb(lang, "x", kanal),
            )

    elif query.data.startswith("tozala_b_ok:") or query.data.startswith("tozala_x_ok:"):
        tip   = "b" if query.data.startswith("tozala_b") else "x"
        kanal = query.data.split(":")[1]
        lang  = context.user_data.get("lang", "cyr")
        ch    = t(lang, CH_KEY.get(kanal, kanal))
        try:
            p = BOT_HOLAT_DIR / (f"buyurtma_{kanal}.json" if tip == "b" else f"xitoy_{kanal}.json")
            if p.exists():
                p.unlink()
            await query.edit_message_text(
                t(lang, "tozala_ok").format(ch=ch),
                parse_mode="Markdown"
            )
        except Exception as e:
            await query.edit_message_text(f"❌ Xato: {e}")

    elif query.data == "yolda_barchasi":
        lang = context.user_data.get("lang", "cyr")
        await query.edit_message_text("⏳ Yuklanmoqda...")
        await yolda_ko_rish(query.message, context, lang)

    elif query.data == "kont_noop":
        pass

    elif query.data == "kont_keldi_ask":
        lang = context.user_data.get("lang", "cyr")
        context.user_data["kutilmoqda"] = ("kont_keldi_sana",)
        context.user_data["screen"] = "keldi_ekran"
        from telegram import ReplyKeyboardMarkup as _RKM
        orqaga_kb = _RKM([[t(lang, "back")]], resize_keyboard=True)
        await query.message.reply_text(
            "📅 Sana *yoki* konteyner nomini kiriting:\n"
            "_(Sana: 07.06.2026 — yoki ISO: CRXU1561318)_",
            parse_mode="Markdown",
            reply_markup=orqaga_kb,
        )
        await query.answer()

    elif query.data == "kont_qaytarish_ask":
        lang = context.user_data.get("lang", "cyr")
        context.user_data["kutilmoqda"] = ("kont_qaytarish_sana",)
        context.user_data["screen"] = "keldi_ekran"
        from telegram import ReplyKeyboardMarkup as _RKM
        orqaga_kb = _RKM([[t(lang, "back")]], resize_keyboard=True)
        await query.message.reply_text(
            "📅 Sana *yoki* konteyner nomini kiriting:\n"
            "_(Sana: 07.06.2026 — yoki ISO: CRXU1561318)_",
            parse_mode="Markdown",
            reply_markup=orqaga_kb,
        )
        await query.answer()

    elif query.data.startswith("kont_bir_keldi:"):
        fname = query.data[len("kont_bir_keldi:"):].split("|", 1)[0]
        old_path = XITOY_PARSED_DIR / fname
        if not old_path.exists():
            await query.answer("Fayl topilmadi.", show_alert=True)
        else:
            iso    = old_path.stem.rsplit("_", 1)[0]
            sana_f = old_path.stem.rsplit("_", 1)[-1]

            # ── Rasmni ALBATTA belgilashdan oldin tayyorlaymiz — aks holda
            # main.py qayta hisoblagach "Yuklangan sana" "Архив" ga almashadi ──
            rasm = generate_kelgan_rasm(iso)

            old_path.rename(XITOY_PARSED_DIR / f"{old_path.stem}_D.xlsx")
            _main_py_ishga_tushir()
            await query.answer(f"✅ {iso} — KELDI!", show_alert=False)
            try:
                await query.message.edit_text(
                    f"✅ *{iso}* ({sana_f}) — KELDI ga o'zgartirildi!",
                    parse_mode="Markdown"
                )
            except Exception:
                pass
            context.user_data["screen"] = "keldi_menu"

            # ── Kelgan konteyner rasmi tayyor — guruhga yuborish so'raymiz ──
            if rasm:
                sent = await query.message.reply_photo(
                    photo=rasm,
                    caption=(
                        f"🖼 *{iso}* — kelgan yuklar ro'yxati.\n"
                        "Guruhga yuborilsinmi?"
                    ),
                    parse_mode="Markdown",
                    reply_markup=InlineKeyboardMarkup([
                        [InlineKeyboardButton("📤 Guruhga jo'natish", callback_data=f"kg_send:{iso}")],
                        [InlineKeyboardButton("❌ Bekor qilish", callback_data=f"kg_cancel:{iso}")],
                    ]),
                )
                context.user_data["kg_pending"] = {
                    "iso": iso,
                    "file_id": sent.photo[-1].file_id,
                }
            else:
                await kont_holat_royhat(query.message, context)

    elif query.data.startswith("kont_bir_qayt:"):
        fname = query.data[len("kont_bir_qayt:"):].split("|", 1)[0]
        old_path = XITOY_PARSED_DIR / fname
        if not old_path.exists():
            await query.answer("Fayl topilmadi.", show_alert=True)
        else:
            stem_no_d = old_path.stem[:-2]
            iso    = stem_no_d.rsplit("_", 1)[0]
            sana_f = stem_no_d.rsplit("_", 1)[-1]
            old_path.rename(XITOY_PARSED_DIR / f"{stem_no_d}.xlsx")
            _main_py_ishga_tushir()
            await query.answer(f"🚢 {iso} — Yo'lda!", show_alert=False)
            try:
                await query.message.edit_text(
                    f"🚢 *{iso}* ({sana_f}) — ЙЎЛДА ga qaytarildi!",
                    parse_mode="Markdown"
                )
            except Exception:
                pass
            context.user_data["screen"] = "keldi_menu"
            await kont_holat_royhat(query.message, context)

    elif query.data.startswith("kont_bekor:"):
        await query.edit_message_text("❌ Bekor qilindi.")

    elif query.data.startswith("kg_send:"):
        iso     = query.data.split(":", 1)[1]
        pending = context.user_data.get("kg_pending")
        if not pending or pending.get("iso") != iso:
            await query.answer("Bu so'rov eskirgan, qayta urinib ko'ring.", show_alert=True)
        else:
            context.user_data["kutilmoqda"] = ("kg_caption", iso, pending["file_id"])
            context.user_data["screen"] = "keldi_menu"
            lang = context.user_data.get("lang", "cyr")
            from telegram import ReplyKeyboardMarkup as _RKM
            orqaga_kb = _RKM([[t(lang, "back")]], resize_keyboard=True)
            await query.answer()
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            await query.message.reply_text(
                "✏️ Rasm ostiga yoziladigan matnni kiriting:",
                reply_markup=orqaga_kb,
            )

    elif query.data.startswith("kg_cancel:"):
        context.user_data.pop("kg_pending", None)
        context.user_data.pop("kutilmoqda", None)
        await query.answer("Bekor qilindi.")
        try:
            await query.edit_message_reply_markup(reply_markup=None)
        except Exception:
            pass
        context.user_data["screen"] = "keldi_menu"
        await kont_holat_royhat(query.message, context)

    elif query.data.startswith("kg_tgl:"):
        payload = query.data[len("kg_tgl:"):]
        fname, _, mk_query_key = payload.partition("|")
        sel_data = context.user_data.get("kg_multi_sel")
        if not sel_data or sel_data.get("query_key") != mk_query_key:
            sel_data = {"query_key": mk_query_key, "selected": set()}
        selected = sel_data["selected"]
        if fname in selected:
            selected.discard(fname)
        else:
            selected.add(fname)
        context.user_data["kg_multi_sel"] = {"query_key": mk_query_key, "selected": selected}
        fayllar = _yolda_fayllar_topish(mk_query_key)
        try:
            await query.edit_message_reply_markup(
                reply_markup=_kg_multi_kb(fayllar, mk_query_key, selected)
            )
        except Exception:
            pass
        await query.answer()

    elif query.data.startswith("kg_confirm_multi:"):
        mk_query_key = query.data.split(":", 1)[1]
        sel_data = context.user_data.get("kg_multi_sel") or {}
        selected = sel_data.get("selected") or set()
        if sel_data.get("query_key") != mk_query_key or not selected:
            await query.answer("Hech narsa tanlanmagan.", show_alert=True)
        else:
            await query.answer("Ishlanmoqda...")
            isos = []
            rasmlar = []   # (iso, BytesIO) — belgilashdan OLDIN tayyorlanadi,
                            # aks holda main.py "Yuklangan sana"ni "Архив" qilib qo'yadi
            for fname in list(selected):
                old_path = XITOY_PARSED_DIR / fname
                if old_path.exists():
                    iso = old_path.stem.rsplit("_", 1)[0]
                    rasm = generate_kelgan_rasm(iso)
                    old_path.rename(XITOY_PARSED_DIR / f"{old_path.stem}_D.xlsx")
                    isos.append(iso)
                    if rasm:
                        rasmlar.append((iso, rasm))
            context.user_data.pop("kg_multi_sel", None)
            context.user_data["screen"] = "keldi_menu"

            if not isos:
                await query.message.reply_text("❌ Tanlangan fayllar topilmadi.")
                await kont_holat_royhat(query.message, context)
            else:
                _main_py_ishga_tushir()
                try:
                    await query.edit_message_text(
                        f"✅ {len(isos)} ta konteyner KELDI ga o'zgartirildi:\n" +
                        "\n".join(f"• {x}" for x in isos)
                    )
                except Exception:
                    pass

                media_items = [InputMediaPhoto(media=rasm) for _, rasm in rasmlar]
                ok_isos = [iso for iso, _ in rasmlar]

                if not media_items:
                    await query.message.reply_text("⚠️ Rasm yaratib bo'lmadi.")
                    await kont_holat_royhat(query.message, context)
                else:
                    sent_msgs = await context.bot.send_media_group(
                        chat_id=query.message.chat_id,
                        media=media_items,
                    )
                    file_ids = [m.photo[-1].file_id for m in sent_msgs if m.photo]
                    context.user_data["kg_pending_multi"] = {
                        "isos": ok_isos,
                        "file_ids": file_ids,
                    }
                    await query.message.reply_text(
                        f"🖼 {len(ok_isos)} ta konteyner rasmi tayyor.\nGuruhga yuborilsinmi?",
                        reply_markup=InlineKeyboardMarkup([
                            [InlineKeyboardButton("📤 Guruhga jo'natish", callback_data="kg_send_multi")],
                            [InlineKeyboardButton("❌ Bekor qilish", callback_data="kg_cancel_multi")],
                        ]),
                    )

    elif query.data == "kg_send_multi":
        pending = context.user_data.get("kg_pending_multi")
        if not pending or not pending.get("file_ids"):
            await query.answer("Bu so'rov eskirgan.", show_alert=True)
        else:
            context.user_data["kutilmoqda"] = ("kg_caption_multi",)
            context.user_data["screen"] = "keldi_menu"
            lang = context.user_data.get("lang", "cyr")
            from telegram import ReplyKeyboardMarkup as _RKM
            orqaga_kb = _RKM([[t(lang, "back")]], resize_keyboard=True)
            await query.answer()
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            await query.message.reply_text(
                "✏️ Rasmlar ostiga yoziladigan matnni kiriting:",
                reply_markup=orqaga_kb,
            )

    elif query.data == "kg_cancel_multi":
        context.user_data.pop("kg_pending_multi", None)
        context.user_data.pop("kutilmoqda", None)
        await query.answer("Bekor qilindi.")
        try:
            await query.edit_message_reply_markup(reply_markup=None)
        except Exception:
            pass
        context.user_data["screen"] = "keldi_menu"
        await kont_holat_royhat(query.message, context)

    elif query.data.startswith("qt_tgl:"):
        payload = query.data[len("qt_tgl:"):]
        fname, _, mk_query_key = payload.partition("|")
        sel_data = context.user_data.get("qt_multi_sel")
        if not sel_data or sel_data.get("query_key") != mk_query_key:
            sel_data = {"query_key": mk_query_key, "selected": set()}
        selected = sel_data["selected"]
        if fname in selected:
            selected.discard(fname)
        else:
            selected.add(fname)
        context.user_data["qt_multi_sel"] = {"query_key": mk_query_key, "selected": selected}
        fayllar = _keldi_fayllar_topish(mk_query_key)
        try:
            await query.edit_message_reply_markup(
                reply_markup=_qt_multi_kb(fayllar, mk_query_key, selected)
            )
        except Exception:
            pass
        await query.answer()

    elif query.data.startswith("qt_confirm_multi:"):
        mk_query_key = query.data.split(":", 1)[1]
        sel_data = context.user_data.get("qt_multi_sel") or {}
        selected = sel_data.get("selected") or set()
        if sel_data.get("query_key") != mk_query_key or not selected:
            await query.answer("Hech narsa tanlanmagan.", show_alert=True)
        else:
            await query.answer("Ishlanmoqda...")
            isos = []
            for fname in list(selected):
                old_path = XITOY_PARSED_DIR / fname
                if old_path.exists():
                    stem_no_d = old_path.stem[:-2]
                    iso = stem_no_d.rsplit("_", 1)[0]
                    old_path.rename(XITOY_PARSED_DIR / f"{stem_no_d}.xlsx")
                    isos.append(iso)
            context.user_data.pop("qt_multi_sel", None)
            context.user_data["screen"] = "keldi_menu"

            if not isos:
                await query.message.reply_text("❌ Tanlangan fayllar topilmadi.")
            else:
                _main_py_ishga_tushir()
                try:
                    await query.edit_message_text(
                        f"↩️ {len(isos)} ta konteyner ЙЎЛДА ga qaytarildi:\n" +
                        "\n".join(f"• {x}" for x in isos)
                    )
                except Exception:
                    pass
            await kont_holat_royhat(query.message, context)

    elif query.data == "qt_cancel_multi":
        context.user_data.pop("qt_multi_sel", None)
        await query.answer("Bekor qilindi.")
        try:
            await query.edit_message_reply_markup(reply_markup=None)
        except Exception:
            pass
        context.user_data["screen"] = "keldi_menu"
        await kont_holat_royhat(query.message, context)

    elif query.data.startswith("karta_kat:"):
        kat  = query.data.split(":")[1]
        lang = context.user_data.get("lang", "cyr")
        kanal = context.user_data.get("kanal", "asosiy")
        key_map = {
            "truba":  "grafik_truba",  "profil": "grafik_profil",
            "list":   "grafik_list",   "bal":    "grafik_bal",
            "stoyka": "grafik_stoyka", "chas":   "grafik_chas",
            "kuz":    "grafik_kuz",    "shar":   "grafik_shar",
            "sokka":  "grafik_sokka",  "oyna":   "grafik_oyna",
        }
        context.user_data["kutilmoqda"] = ("grafik_qidirish", kanal, kat)
        context.user_data["screen"] = "search_kat"
        await query.edit_message_text(
            t(lang, key_map[kat]), parse_mode="Markdown"
        )

    elif query.data == "karta_umumiy":
        lang  = context.user_data.get("lang", "cyr")
        kanal = context.user_data.get("kanal", "asosiy")
        context.user_data["kutilmoqda"] = ("umumiy_qidirish", kanal)
        context.user_data["screen"] = "search_kat"
        await query.edit_message_text(
            "🔍 Qidirmoqchi bo'lgan tovar nomi yoki o'lchamini yozing "
            "(kategoriya tanlash shart emas):",
            parse_mode="Markdown",
        )

    elif query.data.startswith("karta_tovar:"):
        idx  = int(query.data.split(":")[1])
        lang = context.user_data.get("lang", "cyr")
        tovs = context.user_data.pop("grafik_natijalar", [])
        kanal = context.user_data.get("kanal", "asosiy")
        kut   = context.user_data.get("kutilmoqda")
        kat   = kut[2] if isinstance(kut, tuple) and len(kut) > 2 else "truba"
        if idx < len(tovs):
            await query.edit_message_text("⏳ Karta tayyorlanmoqda...")
            await grafik_ko_rsatish(query.message, tovs[idx], kanal, kat)
        else:
            await query.edit_message_text("❌ Xato — qayta urinib ko'ring.")

    elif query.data == "tozala_no":
        lang = context.user_data.get("lang", "cyr")
        await query.edit_message_text(t(lang, "tozala_bekor"))

    # ── Zakaz 2-bosqich tasdiqlash ────────────────────────────
    elif query.data.startswith("zakaz_ok:"):
        kanal   = query.data.split(":")[1]
        lang    = context.user_data.get("lang", "cyr")
        user_id = query.from_user.id
        pending = context.user_data.pop("pending_zakaz", None)

        # RAM da yo'q bo'lsa (bot qayta ishga tushgan) — diskdan o'qiymiz
        if pending is None:
            disk_items = pending_yuklash(kanal, user_id)
            if disk_items is not None:
                pending = (kanal, disk_items)

        if pending and pending[0] == kanal:
            try:
                buyurtma_saqlash(kanal, pending[1])
            except Exception as e:
                import traceback
                tb = traceback.format_exc()
                await query.edit_message_text(f"❌ buyurtma_saqlash xato:\n```\n{tb[-1200:]}\n```", parse_mode="Markdown")
                return
            pending_tozala(kanal, user_id)
            await query.edit_message_text(
                t(lang, "tasdiq_ok").format(n=len(pending[1])),
                parse_mode="Markdown"
            )
        else:
            await query.edit_message_text(
                f"❌ Sessiya eskirdi.\npending={pending!r}\nkanal={kanal!r}"
            )

    elif query.data == "zakaz_no":
        lang    = context.user_data.get("lang", "cyr")
        user_id = query.from_user.id
        # Avval qaysi kanalga tegishli ekanini bilishimiz kerak
        # Agar pending_zakaz RAMda bor bo'lsa — undan olamiz, yo'qsa barcha kanallarni tozalaymiz
        pending = context.user_data.pop("pending_zakaz", None)
        if pending:
            pending_tozala(pending[0], user_id)
        else:
            for _k in ["asosiy", "cex", "osh"]:
                pending_tozala(_k, user_id)
        await query.edit_message_text(
            t(lang, "zakaz_bekor_msg"), parse_mode="Markdown"
        )

    # ── Yo'lga konteyner qo'shish tasdiq ─────────────────────────────────────
    elif query.data.startswith("kont:"):
        lang     = context.user_data.get("lang", "cyr")
        decision = query.data.split(":")[1]
        if decision == "yoq":
            context.user_data.pop("kont_yangilar", None)
            context.user_data.pop("kutilmoqda", None)
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            await query.message.reply_text(t(lang, "kont_tasdiq_yoq"))
        elif decision == "ha":
            yangilar = context.user_data.pop("kont_yangilar", [])
            context.user_data.pop("kutilmoqda", None)
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            if not yangilar:
                await query.message.reply_text(t(lang, "kont_barchasi_bor"))
                return
            from konteyner_qosh import konteyner_xlsx_yarat
            n_ok = 0
            qoshilganlar = []
            for kont in yangilar:
                try:
                    # XITOY_PARSED_DIR — main.py aynan shu papkani o'qib,
                    # "Yo'ldagi yuklar" hisobotiga qo'shadi (KONT_DIR emas!)
                    konteyner_xlsx_yarat(kont, XITOY_PARSED_DIR)
                    n_ok += 1
                    qoshilganlar.append(kont)
                except Exception as e:
                    logger.error(f"konteyner xlsx xato {kont['iso']}: {e}")
            if qoshilganlar:
                # Doimiy tarixga (ISO+sana) yozamiz — fayl keyin o'chirilsa ham
                # aynan shu yetkazib berish qayta "yangi" deb qo'shilib ketmaydi
                # (lekin xuddi shu ISO boshqa sanada YANGI yuk bilan kelsa —
                # bemalol qo'shiladi, chunki konteyner raqamlari qayta ishlatiladi)
                konteyner_tarix_qoshish(qoshilganlar)
            if n_ok:
                _main_py_ishga_tushir()
            await query.message.reply_text(
                t(lang, "kont_qoshildi").format(n=n_ok),
                parse_mode="Markdown"
            )


async def text_keldi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Barcha matn xabarlarini boshqaradi — navigatsiya va qidiruv."""
    uid = update.message.from_user.id
    if not kirish_ruxsati(uid):
        await update.message.reply_text("⛔ Kirish huquqi yo'q.")
        return
    context.user_data["user_id"] = uid  # _is_admin uchun
    lang   = context.user_data.get("lang", "cyr")
    screen = context.user_data.get("screen", "main")
    text   = update.message.text.strip()
    msg    = update.message
    loop   = asyncio.get_event_loop()
    import logging as _log
    _log.getLogger(__name__).info("text_keldi: lang=%s screen=%s text=%r", lang, screen, text)

    # Fayl kutilayotgan bo'lsa:
    # xitoy_ostatka_fayl holatida menyu tugmasi kelsa — normal ishlat (kutilmoqda tozala)
    # Tasodifiy matn kelsa — eslatma, holatni saqlat
    kut = context.user_data.get("kutilmoqda")

    # ── Konteyner qidiruv holatlari (sana yoki nom bo'yicha) ─────────────────
    if isinstance(kut, tuple) and kut[0] in ("kont_keldi_sana", "kont_qaytarish_sana"):
        _back_texts = {t("cyr", "back"), t("lat", "back"), "⬅️ Орқага", "⬅️ Orqaga", "← Orqaga"}
        if text in _back_texts or bool(get_action(lang, screen, text)):
            context.user_data.pop("kutilmoqda", None)
        else:
            import re as _re
            rejim = kut[0]
            is_keldi_rejim = (rejim == "kont_keldi_sana")

            if _re.match(r"^\d{2}\.\d{2}\.\d{4}$", text):
                query_key = f"sana:{text}"
            else:
                query_key = f"nom:{text.strip().upper()}"

            context.user_data.pop("kutilmoqda", None)
            await _kont_list_yuborish(msg, query_key, is_keldi_rejim,
                                       reply=True, context=context)
            return
    # ─────────────────────────────────────────────────────────────────────────

    # ── "Kelgan yuklar" guruhiga yuborish uchun matn (caption) kutilmoqda ──────
    if isinstance(kut, tuple) and kut[0] == "kg_caption":
        _back_texts = {t("cyr", "back"), t("lat", "back"), "⬅️ Орқага", "⬅️ Orqaga", "← Orqaga"}
        _, kg_iso, kg_file_id = kut
        if text in _back_texts or bool(get_action(lang, screen, text)):
            context.user_data.pop("kutilmoqda", None)
            context.user_data.pop("kg_pending", None)
        else:
            context.user_data.pop("kutilmoqda", None)
            context.user_data.pop("kg_pending", None)
            await _kg_yuborish_guruhga(msg, context, kg_iso, kg_file_id, text)
            return
    # ─────────────────────────────────────────────────────────────────────────

    # ── Bir nechta konteyner (albom) uchun caption kutilmoqda ──────────────────
    if isinstance(kut, tuple) and kut[0] == "kg_caption_multi":
        _back_texts = {t("cyr", "back"), t("lat", "back"), "⬅️ Орқага", "⬅️ Orqaga", "← Orqaga"}
        pending = context.user_data.get("kg_pending_multi") or {}
        if text in _back_texts or bool(get_action(lang, screen, text)):
            context.user_data.pop("kutilmoqda", None)
            context.user_data.pop("kg_pending_multi", None)
        else:
            context.user_data.pop("kutilmoqda", None)
            context.user_data.pop("kg_pending_multi", None)
            await _kg_yuborish_guruhga_multi(
                msg, context,
                pending.get("isos", []), pending.get("file_ids", []), text
            )
            return
    # ─────────────────────────────────────────────────────────────────────────

    _KAT_MATN = {
        "truba":  "grafik_truba",  "profil": "grafik_profil",
        "list":   "grafik_list",   "bal":    "grafik_bal",
        "stoyka": "grafik_stoyka", "chas":   "grafik_chas",
        "kuz":    "grafik_kuz",    "shar":   "grafik_shar",
        "sokka":  "grafik_sokka",  "oyna":   "grafik_oyna",
    }
    if isinstance(kut, tuple) and kut[0] == "grafik_qidirish":
        kanal = kut[1]
        kat   = kut[2]
        # Orqaga yoki menyu tugmasi kelsa — holatni tozala
        if text == t(lang, "back") or get_action(lang, screen, text):
            if text == t(lang, "back"):
                if context.user_data.pop("grafik_back", False):
                    # 2-Orqaga: search_kat → kategoriya tanlash
                    context.user_data.pop("kutilmoqda", None)
                    context.user_data.pop("grafik_natijalar", None)
                    await go_screen(msg, context, "search")
                    await msg.reply_text(t(lang, "karta_kat_sor"), reply_markup=grafik_kat_ikb())
                else:
                    # 1-Orqaga: natijadan → yozish ekrani (kutilmoqda saqlanadi)
                    context.user_data.pop("grafik_natijalar", None)
                    context.user_data["grafik_back"] = True
                    context.user_data["screen"] = "search_kat"
                    matn_key = _KAT_MATN.get(kat, "grafik_list")
                    await msg.reply_text(t(lang, matn_key), parse_mode="Markdown",
                                         reply_markup=search_kb(lang))
                return
            # Boshqa menyu tugmasi
            context.user_data.pop("kutilmoqda", None)
            context.user_data.pop("grafik_natijalar", None)
            context.user_data.pop("grafik_back", None)
        else:
            # Qidiruv so'rovi — yangi qidiruv, grafik_back ni tozala
            context.user_data.pop("grafik_back", None)
            df = grafik_qidirish(text, kat, kanal)
            if df.empty:
                await msg.reply_text(t(lang, "grafik_topilmadi"), parse_mode="Markdown")
                return
            tovs = df["Товар"].tolist()
            if len(tovs) == 1:
                # kutilmoqda ni saqlaymiz — davomiy qidiruv uchun
                await msg.reply_text("⏳ Grafik chizilmoqda...")
                await grafik_ko_rsatish(msg, tovs[0], kanal, kat)
            else:
                context.user_data["grafik_natijalar"] = tovs
                await msg.reply_text(
                    t(lang, "grafik_tanlang"),
                    parse_mode="Markdown",
                    reply_markup=grafik_tovar_ikb(tovs),
                )
            return

    # ── Umumiy (kategoriyasiz) qidiruv ───────────────────────────────────────
    if isinstance(kut, tuple) and kut[0] == "umumiy_qidirish":
        kanal = kut[1]
        if text == t(lang, "back"):
            context.user_data.pop("kutilmoqda", None)
            await go_screen(msg, context, "search")
            await msg.reply_text(t(lang, "karta_kat_sor"), reply_markup=grafik_kat_ikb())
            return
        if get_action(lang, screen, text):
            # Boshqa menyu tugmasi — holatni tozalab, odatiy amalga o'tamiz
            context.user_data.pop("kutilmoqda", None)
        else:
            df = qidiruv_olish(text, kanal)
            matn = qidiruv_text(text, df, lang)
            await msg.reply_text(matn, parse_mode="Markdown")
            return  # kutilmoqda saqlanadi — davomli qidiruv uchun

    if isinstance(kut, tuple):
        if kut[0] in ("xitoy_tp", "xitoy_list", "xitoy_ostatka_fayl", "xitoy_fayl"):
            is_back    = (text == t(lang, "back"))
            action_now = get_action(lang, screen, text)
            # "hisoblash" tugmasi — kutilmoqdani tozalama, davom et (handler o'zi tozalaydi)
            if action_now == "hisoblash":
                pass  # quyida hisoblash handleri ishlaydi
            elif is_back or bool(action_now):
                # Boshqa menyu tugmasi — holatni tozala
                context.user_data.pop("kutilmoqda", None)
                context.user_data.pop("xitoy_akkum", None)
            else:
                # Tasodifiy matn — eslatma, holatni saqlat
                await msg.reply_text(
                    "📎 Excel fayl kutilmoqda. Faylni yuboring yoki *▶️ Hisoblash* tugmasini bosing.",
                    parse_mode="Markdown",
                )
                return
        else:
            context.user_data.pop("kutilmoqda", None)

    # Orqaga (universal — har qanday ekranda ishlaydi)
    _back_set = {t("cyr", "back"), t("lat", "back"), "⬅️ Орқага", "⬅️ Orqaga", "← Orqaga"}
    if text in _back_set:
        context.user_data.pop("kutilmoqda", None)
        context.user_data.pop("grafik_natijalar", None)
        # keldi_ekran dan Orqaga → inline menyu ko'rsatish, keldi_menu ga o'tish
        if screen == "keldi_ekran":
            context.user_data["screen"] = "keldi_menu"
            await kont_holat_royhat(msg, context)
            return
        parent = BACK_MAP.get(screen, "main")
        await go_screen(msg, context, parent)
        if parent == "search":
            await msg.reply_text(t(lang, "karta_kat_sor"), reply_markup=grafik_kat_ikb())
        return

    # Amal topish
    action = get_action(lang, screen, text)

    if action is None:
        # search ekranida (kategoriya tanlash) — matn yozilsa eslatma
        if screen == "search":
            await msg.reply_text(t(lang, "karta_kat_sor"), reply_markup=grafik_kat_ikb())
            return
        # Tanilmagan tugma — joriy ekranni qayta ko'rsatamiz
        await go_screen(msg, context, screen)
        return

    # Kanal ekraniga o'tish
    if isinstance(action, tuple):
        next_screen, kanal = action
        await go_screen(msg, context, next_screen, kanal=kanal)
        # load_channel ga o'tganda darhol fayl kutish holatiga tush
        if next_screen == "load_channel":
            context.user_data["kutilmoqda"]    = ("xitoy_fayl", kanal)
            context.user_data["xitoy_akkum"]   = {}   # akkumulyator reset
            context.user_data["ombor_akkum"]   = {}   # ombor akkumulyator reset
            context.user_data["vazn_akkum"]    = {}   # vazn akkumulyator reset
        return

    # Oddiy navigatsiya
    if action in ("order","load","settings","search","main"):
        await go_screen(msg, context, action)
        return

    # Buyurtma amallari
    kanal = context.user_data.get("kanal","asosiy")

    if action == "excel":
        # Avval Xitoy ostatka haqida so'raymiz.
        # Mavjud JSON bor bo'lsa — uni ishlatish/yangilash/hisobsiz tanlash.
        mavjud = xitoy_yuklash(kanal)
        ch     = t(lang, CH_KEY[kanal])

        # Tasdiqlangan buyurtma ma'lumoti (har ikkala holatda ham ko'rsatiladi)
        b_data = buyurtma_yuklash(kanal)
        if b_data and b_data.get("buyurtmalar"):
            b_n    = len(b_data["buyurtmalar"])
            b_sana = b_data.get("sana", "?")
            b_line_cyr = f"\n\n📋 *Тасдиқланган буюртма:* {b_n} та\n🗓 Буюртма санаси: *{b_sana}*"
            b_line_lat = f"\n\n📋 *Tasdiqlangan buyurtma:* {b_n} ta\n🗓 Buyurtma sanasi: *{b_sana}*"
        else:
            b_line_cyr = "\n\n📋 *Тасдиқланган буюртма:* йўқ"
            b_line_lat = "\n\n📋 *Tasdiqlangan buyurtma:* yo'q"
        b_line = b_line_cyr if lang == "cyr" else b_line_lat

        if mavjud is not None and mavjud.get("tovarlar"):
            # Xitoy ostatka bor
            n    = len(mavjud["tovarlar"])
            sana = mavjud.get("sana", "?")
            await msg.reply_text(
                t(lang, "xitoy_mavjud").format(ch=ch, n=n, sana=sana) + b_line,
                parse_mode="Markdown",
                reply_markup=xitoy_mavjud_ikb(lang, kanal),
            )
        elif mavjud is not None:
            # xitoy_saqlash({}) chaqirilgan — "shundaylicha" tanlangan edi
            # Yana shundaylicha yoki yangi xitoy yuklash
            hdr_cyr = f"📦 *{ch}*\n\n⚠️ Хитой охирги маълумот: *Захирасиз* юборилган."
            hdr_lat = f"📦 *{ch}*\n\n⚠️ Xitoy oxirgi ma'lumot: *Zahirasiz* yuborilgan."
            hdr = hdr_cyr if lang == "cyr" else hdr_lat
            await msg.reply_text(
                hdr + b_line,
                parse_mode="Markdown",
                reply_markup=xitoy_mavjud_ikb(lang, kanal),
            )
        else:
            # Birinchi marta — hech qanday xitoy ma'lumot yo'q
            await msg.reply_text(
                t(lang, "xitoy_sorash"),
                parse_mode="Markdown",
                reply_markup=xitoy_sorash_ikb(lang, kanal),
            )

    elif action == "tasdiq":
        context.user_data["kutilmoqda"] = ("buyurtma_tasdiq", kanal)
        ch = t(lang, CH_KEY[kanal])
        await msg.reply_text(
            t(lang,"tasdiq_prompt").format(ch=ch),
            parse_mode="Markdown",
        )

    elif action == "upload":
        context.user_data["kutilmoqda"]  = ("xitoy_fayl", kanal)
        context.user_data["xitoy_akkum"] = {}
        context.user_data["ombor_akkum"] = {}
        context.user_data["vazn_akkum"]  = {}
        context.user_data["kanal"]       = kanal
        ch = t(lang, CH_KEY[kanal])
        await msg.reply_text(
            t(lang,"load_ch_title").format(ch=ch),
            parse_mode="Markdown",
        )

    elif action == "hisoblash":
        akkum       = context.user_data.get("xitoy_akkum", {})
        ombor_akkum = context.user_data.get("ombor_akkum", {})
        if not akkum and not ombor_akkum:
            await msg.reply_text(
                "⚠️ Hali hech qanday fayl yuborilmadi. Avval xitoy Excel faylini yuboring.",
                parse_mode="Markdown",
            )
            return
        async with yuklash_animatsiya(
            msg, context,
            text_cyr="Юклатиш режаси тайёрланмоқда",
            text_lat="Yuklatish rejasi tayyorlanmoqda",
        ):
            try:
                import importlib.util as _ilu
                _nd = Path(__file__).resolve().parent
                for _dep in ("vazn_hisobla", "Yuklama_optimal", "yuklatish_rejasi"):
                    if _dep not in sys.modules:
                        _sp = _ilu.spec_from_file_location(_dep, _nd / f"{_dep}.py")
                        _m  = _ilu.module_from_spec(_sp)
                        sys.modules[_dep] = _m
                        _sp.loader.exec_module(_m)
                main_with_data = sys.modules["yuklatish_rejasi"].main_with_data
                ombor_akkum = context.user_data.get("ombor_akkum", {})
                vazn_akkum  = context.user_data.get("vazn_akkum", {})
                xlsx_path = main_with_data(kanal, ombor_akkum, xitoy_vazn=vazn_akkum)
            except Exception as e:
                await msg.reply_text(
                    f"❌ Yuklatish rejasi yaratishda xato:\n{str(e)[:300]}"
                )
                return
        context.user_data.pop("xitoy_akkum", None)
        context.user_data.pop("ombor_akkum", None)
        context.user_data.pop("vazn_akkum", None)
        context.user_data.pop("kutilmoqda", None)

        # STATS:konteyner|yuklangan_kg|yuklangan_xil|qolgan_xil|qolgan_kg|path
        stats_caption = ""
        if xlsx_path and xlsx_path.startswith("STATS:"):
            parts_s = xlsx_path.split("|")
            try:
                n_kont   = parts_s[0].split(":")[1]
                yuk_kg   = float(parts_s[1])
                yuk_xil  = parts_s[2]
                qol_xil  = int(parts_s[3])
                qol_kg   = float(parts_s[4])
                xlsx_path = parts_s[5] if len(parts_s) > 5 else ""
                yuk_t = yuk_kg / 1000
                stats_caption = (
                    f"✅ *{n_kont} konteyner* yuklandi — "
                    f"*{yuk_xil} xil tovar*, *{yuk_t:.1f} t*\n"
                )
                if qol_xil:
                    qol_t = qol_kg / 1000
                    stats_caption += (
                        f"📋 Yuklanmadi: *{qol_xil} xil tovar* "
                        f"({qol_t:.1f} t ostatkada)\n"
                    )
            except Exception:
                xlsx_path = xlsx_path.split("|")[-1] if "|" in xlsx_path else xlsx_path
                stats_caption = ""

        if not xlsx_path or not xlsx_path.endswith(".xlsx"):
            if xlsx_path == "KERAK_YOQ":
                await msg.reply_text(
                    "✅ Power BI da hozircha kamomati bor tovar yo'q (Кам = 0 barcha tovarlar uchun).\n"
                    "Bu to'g'ri natija bo'lishi mumkin — yoki Power BI ma'lumotlari yangilanmagan."
                )
            elif xlsx_path == "OMBOR_BOʻSH":
                await msg.reply_text("⚠️ Xitoy omborida tovar topilmadi (miqdorlar 0).")
            elif xlsx_path and xlsx_path.startswith("MOS_YOQ|"):
                parts = xlsx_path.split("|")
                n_kerak = parts[1] if len(parts) > 1 else "?"
                n_mavjud = parts[2] if len(parts) > 2 else "?"
                namuna = (parts[3] if len(parts) > 3 else "").replace(";;", "\n  • ")
                await msg.reply_text(
                    f"⚠️ *Tovar nomlari mos kelmadi!*\n\n"
                    f"Kamomati bor: *{n_kerak}* ta tovar\n"
                    f"Xitoy faylida: *{n_mavjud}* ta tovar\n"
                    f"Mos kelgani: *0* ta\n\n"
                    f"Xitoy faylidagi nomlar:\n  • {namuna}\n\n"
                    f"Bu nomlar inventar (Power BI) dagi nomlar bilan bir xil bo'lishi kerak.",
                    parse_mode="Markdown",
                )
            else:
                await msg.reply_text("⚠️ Yuklatish rejasi bo'sh yoki tovar mos kelmadi.")
            return
        try:
            with open(xlsx_path, "rb") as f:
                await msg.reply_document(
                    document=f,
                    filename=Path(xlsx_path).name,
                    caption=(
                        stats_caption +
                        f"📦 *Yuklatish Rejasi — {kanal.upper()}*\n"
                        f"Xitoy omborida tayyor tovarlardan optimal yuklatish plani."
                    ),
                    parse_mode="Markdown",
                )
        except Exception as e:
            await msg.reply_text(
                f"Excel tayyor, lekin yuborishda xato: {str(e)[:150]}\nFayl: {xlsx_path}"
            )
        # Hisoblash tugagach — load_channel ekraniga qayt (foydalanuvchi qayta hisoblashi mumkin)
        await go_screen(msg, context, "load_channel", kanal=kanal)

    elif action == "karta":
        context.user_data["screen"] = "search"
        # reply keyboard ni o'rnatish (Orqaga tugmasi)
        await msg.reply_text("🔍", reply_markup=search_kb(lang))
        # kategoriya inline keyboard
        await msg.reply_text(t(lang, "karta_kat_sor"), reply_markup=grafik_kat_ikb())

    elif action == "yolda":
        pass  # Eski tugma — hozircha bo'sh

    elif action == "konteyner":
        await go_screen(msg, context, "konteyner")

    elif action == "keldi_belgi":
        context.user_data["screen"] = "keldi_menu"
        await kont_holat_royhat(msg, context, change_kb=True)

    elif action == "yangilash":
        loop = asyncio.get_event_loop()
        async with yuklash_animatsiya(
            msg, context,
            text_cyr="Маълумотлар янгиланмоқда",
            text_lat="Ma'lumotlar yangilanmoqda",
        ):
            try:
                # DIQQAT (2026-07-08 qo'shildi): Huzayfa kunlik qoldiqni faqat
                # o'z ish kompyuterida tarix/ ga qo'shib, git'ga push qiladi —
                # serverga jismonan/AnyDesk orqali kirmasdan, Telegram'dan
                # "Yangilash" tugmasi orqali serverdagi eng so'nggi ma'lumotni
                # (git pull) tortib olib, shu yerda hisoblashi kerak. Shuning
                # uchun main.py'dan OLDIN avval "git pull" qilinadi. Bu FAQAT
                # ma'lumot (tarix/) yangilanishlari uchun — kod (.py fayllar)
                # o'zgarishlari BUNDAN keyin ham pm2 restart talab qiladi
                # (server_yangilash.bat), chunki Python allaqachon xotiraga
                # yuklangan kodni o'zi qayta yuklay olmaydi.
                git_proc = await asyncio.create_subprocess_exec(
                    "git", "pull", "origin", "main",
                    stdout=asyncio.subprocess.PIPE,
                    stderr=asyncio.subprocess.PIPE,
                    cwd=str(BASE_DIR),
                )
                git_out, git_err = await git_proc.communicate()
                if git_proc.returncode != 0:
                    xato = (git_err or git_out).decode("utf-8", errors="replace")[-300:]
                    await msg.reply_text(
                        t(lang, "yangilash_err").format(xato=f"git pull xato: {xato}")
                    )
                    return

                main_py = str(BASE_DIR / "main.py")
                env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONLEGACYWINDOWSSTDIO": "0"}
                proc = await asyncio.create_subprocess_exec(
                    sys.executable, main_py,
                    stdout=asyncio.subprocess.PIPE,
                    stderr=asyncio.subprocess.PIPE,
                    cwd=str(BASE_DIR),
                    env=env,
                )
                _, stderr = await proc.communicate()
                if proc.returncode == 0:
                    # DIQQAT (2026-07-08 tuzatildi): kesh FAQAT main.py
                    # muvaffaqiyatli tugagandan KEYIN yangilanishi kerak —
                    # avval edi (main.py'dan OLDIN), shuning uchun tugma
                    # bosilgandan keyin ham bot xotirasi ESKI faylni ko'rsatib
                    # turardi (yangi fayl diskka yozilgach ham, keshga
                    # qayta o'qilmasdi, keyingi 5 daqiqalik TTL tugagunicha).
                    await loop.run_in_executor(None, lambda: _xlsx_refresh(force=True))
                    await msg.reply_text(t(lang, "yangilash_ok"))
                else:
                    xato = stderr.decode("utf-8", errors="replace")[-300:]
                    await msg.reply_text(
                        t(lang, "yangilash_err").format(xato=xato)
                    )
            except Exception as e:
                await msg.reply_text(
                    t(lang, "yangilash_err").format(xato=str(e)[:300])
                )

    elif action == "lang_pick":
        await msg.reply_text("🇺🇿 Тилни танланг:", reply_markup=til_ikb())

    elif action == "tozala_buy":
        await msg.reply_text(
            t(lang, "tozala_sor_buy"),
            reply_markup=tozala_kanal_ikb(lang, "b"),
        )

    elif action == "tozala_xitoy":
        await msg.reply_text(
            t(lang, "tozala_sor_xitoy"),
            reply_markup=tozala_kanal_ikb(lang, "x"),
        )

    elif action == "yolga_kont":
        uid = update.message.from_user.id
        if ADMIN_IDS and uid not in ADMIN_IDS:
            await msg.reply_text("❌ Bu funksiya faqat admin uchun.")
            return
        context.user_data["kutilmoqda"] = ("kont_tp", None)
        context.user_data.pop("kont_tp_raw", None)
        await msg.reply_text(t(lang, "kont_tp_kut"), parse_mode="Markdown")

    elif action == "boglanish":
        ikb_kb = boglanish_ikb(lang, SUPPORT_PHONE, SUPPORT_USERNAME)
        lines = []
        if SUPPORT_PHONE:
            lines.append(f'<a href="tel:{SUPPORT_PHONE}">📞 {SUPPORT_PHONE}</a>')
        sarlavha = "👤 <b>Admin bilan bog'lanish:</b>" if lang == "lat" else "👤 <b>Admin билан боғланиш:</b>"
        matn = sarlavha + "\n\n" + "\n".join(lines) if lines else "Bog'lanish ma'lumoti hali kiritilmagan."
        await msg.reply_text(matn, parse_mode="HTML", reply_markup=ikb_kb)

    elif action == "yolda_excel":
        # Admin bilan bir xil — Power BI Excelidan yo'lda holat
        await yolda_ko_rish(msg, context, lang)


async def fayl_keldi(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Xitoy Excel fayllarini qabul qiladi."""
    uid = update.message.from_user.id
    if not kirish_ruxsati(uid):
        await update.message.reply_text("⛔ Kirish huquqi yo'q.")
        return
    context.user_data["user_id"] = uid  # _is_admin uchun
    lang  = context.user_data.get("lang", "cyr")
    msg   = update.message
    kut   = context.user_data.get("kutilmoqda")

    if not (isinstance(kut, tuple) and kut[0] in ("xitoy_fayl", "xitoy_ostatka_fayl", "xitoy_tp", "xitoy_list", "buyurtma_tasdiq", "kont_tp", "kont_list", "kont_tasdiq_fayl")):
        await msg.reply_text("Fayl kutilmayapti. Avval kanal tanlang.")
        return

    kanal = kut[1]
    doc   = msg.document
    if doc is None:
        await msg.reply_text("Hujjat topilmadi.")
        return

    # DIQQAT: bu xabar Telegram'dan faylni YUKLAB OLISHDAN OLDIN
    # yuboriladi — aks holda (avvalgi kod shunday edi) katta fayllarda
    # bot bir necha soniya JIM turib qolar edi (fayl hali serverdan
    # yuklab olinayotgan payt hech qanday javob yo'q edi), va foydalanuvchi
    # botni "osilib qoldi" deb o'ylardi. Har bir holat (kont_tp, kont_list
    # va h.k.) o'zining KEYINGI, aniqroq xabarini pastda alohida yuboradi.
    await msg.reply_text("📎 Fayl qabul qilindi, yuklanmoqda...")

    try:
        tg_file = await doc.get_file()
        bio = BytesIO()
        await tg_file.download_to_memory(bio)
        bio.seek(0)
        raw = bio.read()
    except Exception as e:
        await msg.reply_text(f"Fayl yuklanmadi: {e}")
        return

    # ── Tasdiqlangan buyurtma fayli ──────────────────────────────────────────
    if kut[0] == "buyurtma_tasdiq":
        try:
            ok, xato, items = buyurtma_tekshir(raw, kanal)
        except Exception as e:
            await msg.reply_text(f"❌ Buyurtma tekshirishda xato: {str(e)[:300]}")
            return
        if not ok:
            await msg.reply_text(
                t(lang, "tasdiq_err").format(xato=xato),
                parse_mode="Markdown",
            )
            return
        if not items:
            await msg.reply_text(
                "⚠️ Buyurtma ustunida hech qanday miqdor topilmadi.\n"
                "Excel faylida 'Buyurtma' ustunini to'ldiring va qayta yuboring."
            )
            return
        user_id = update.message.from_user.id
        context.user_data["pending_zakaz"] = (kanal, items)
        try:
            pending_saqlash(kanal, user_id, items)
        except Exception:
            pass
        context.user_data.pop("kutilmoqda", None)
        try:
            preview = zakaz_preview_text(kanal, items, lang)
        except Exception as e:
            preview = f"Oldingi tasdiq: {e}"
        await msg.reply_text(
            preview,
            parse_mode="Markdown",
            reply_markup=zakaz_tasdiq_ikb(lang, kanal),
        )
        return

    # ── Yo'lga konteyner qo'shish: 1/2 — Труба/Профиль装箱单 ────────────────────
    if kut[0] == "kont_tp":
        # DIQQAT: bu yerda hali og'ir tekshirish/parslash YO'Q — faqat
        # xotiraga saqlanadi, haqiqiy o'qish ikkala fayl ham kelgach
        # ("kont_list" bosqichida) bajariladi. Shuning uchun qo'shimcha
        # "tekshirilmoqda" xabari YOLG'ON va ORTIQCHA edi (yuqoridagi umumiy
        # "qabul qilindi" xabari yetarli) — olib tashlandi.
        context.user_data["kont_tp_raw"] = raw
        context.user_data["kutilmoqda"] = ("kont_list", None)
        await msg.reply_text(t(lang, "kont_list_kut"), parse_mode="Markdown")
        return

    # ── Yo'lga konteyner qo'shish: 2/2 — Лист出货清单 → tekshirish uchun draft ──
    if kut[0] == "kont_list":
        from konteyner_qosh import (
            xitoy_yuklar_oqi, draft_excel_yarat, qisqa_xulosa,
            iso_boyicha_yangilarini_ajrat, notanish_soni,
            oxirgi_malum_sana, faqat_sanadan_keyingi,
        )
        await msg.reply_text("⏳ Ikkala fayl o'qilmoqda va solishtirilmoqda, biroz kuting...")
        truba_raw = context.user_data.pop("kont_tp_raw", None)
        if not truba_raw:
            await msg.reply_text(
                "⚠️ 1-fayl (Труба/Профиль) topilmadi — qaytadan boshlang."
            )
            context.user_data.pop("kutilmoqda", None)
            return
        try:
            yuklar = xitoy_yuklar_oqi(truba_raw, raw)
        except Exception as e:
            await msg.reply_text(t(lang, "kont_xato").format(xato=str(e)[:300]), parse_mode="Markdown")
            context.user_data.pop("kutilmoqda", None)
            return
        tarix = konteyner_tarix_olish()
        # 1-QADAM — GLOBAL sana chegarasi: Xitoy faylida (kumulyativ master
        # ro'yxat bo'lgani uchun) ko'p oy oldingi (masalan mart/aprel)
        # yozuvlar ham bor bo'lishi mumkin. Bular ILGARI "ISO hech qachon
        # uchramagan bo'lsa — sanasidan qat'iy nazar yangi" qoidasi orqali
        # noto'g'ri qoshilib ketardi (masalan yangi payqalgan mashina-raqam
        # psevdo-ID'lari yoki tuzatilgan ISO'lar). Endi tizimdagi ENG OXIRGI
        # ma'lum yuklangan sanadan KEYINGI (qat'iy katta) yozuvlar
        # QOLDIRILADI — bundan eskisi butunlay o'tkazib yuboriladi.
        oxirgi = oxirgi_malum_sana(XITOY_PARSED_DIR, tarix)
        soni_oldin = len(yuklar)
        yuklar = faqat_sanadan_keyingi(yuklar, oxirgi)
        eskisi_soni = soni_oldin - len(yuklar)
        if oxirgi and eskisi_soni:
            await msg.reply_text(
                f"ℹ️ Tizimdagi eng oxirgi ma'lum sana: {oxirgi.strftime('%d.%m.%Y')}. "
                f"Shundan OLDINGI (eski) {eskisi_soni} ta yozuv Xitoy faylida "
                f"topildi va o'tkazib yuborildi."
            )
        # 2-QADAM — shundan keyin, qolgan (yangi sanadagi) yozuvlar orasida
        # HAR BIR ISO alohida tekshiriladi: shu ISO ushbu sanadan avvalroq
        # allaqachon xuddi shu (yoki keyingi) sana bilan qayd etilgan bo'lsa
        # — qayta o'tkazib yuboriladi.
        yangilar = iso_boyicha_yangilarini_ajrat(yuklar, XITOY_PARSED_DIR, tarix)
        if not yangilar:
            await msg.reply_text(t(lang, "kont_barchasi_bor"), parse_mode="Markdown")
            context.user_data.pop("kutilmoqda", None)
            return
        context.user_data["kont_yangilar"] = yangilar
        context.user_data["kutilmoqda"] = ("kont_tasdiq_fayl", None)
        # Tasdiqlashdan oldin QISQA xulosa (har bir konteyner — birlashganmi,
        # faqat Труба/Лист, tonnaji, inventarda mos kelmagan tovar bormi) +
        # to'liq tafsilotli Excel + tasdiqlash tugmasi.
        await msg.reply_text(qisqa_xulosa(yangilar), parse_mode="Markdown")
        bio = draft_excel_yarat(yangilar)
        await msg.reply_document(
            document=bio, filename="Yangi_konteynerlar.xlsx",
            caption=(
                "📋 Har bir konteynerning to'liq tovar ro'yxati shu faylda. "
                "Xato bo'lsa tahrirlab qayta shu yerga yuboring. "
                "Hammasi to'g'ri bo'lsa — pastdagi tugmani bosing."
            ),
        )
        n_nomos = notanish_soni(yangilar)
        if n_nomos:
            await msg.reply_text(
                f"⚠️ *DIQQAT: {n_nomos} ta tovar hali ham inventarda "
                f"NOTANISH* (faylda ⚠️ belgi bilan ajratilgan). Iltimos, "
                f"yuqoridagi Excel faylni oching, shu qatorlarni TO'G'IRLAB "
                f"qayta shu yerga yuboring — aks holda ular xato/noaniq nom "
                f"bilan saqlanib qoladi.",
                parse_mode="Markdown",
            )
            await msg.reply_text(
                "✅ Baribir tasdiqlaysizmi?",
                reply_markup=kont_tasdiq_ikb(lang),
            )
        else:
            await msg.reply_text(
                "✅ Tasdiqlaysizmi?",
                reply_markup=kont_tasdiq_ikb(lang),
            )
        return

    # ── Yo'lga konteyner qo'shish: tasdiqlashdan oldin tahrirlangan fayl ──────
    if kut[0] == "kont_tasdiq_fayl":
        from konteyner_qosh import (
            draft_excel_oqi, draft_excel_yarat, qisqa_xulosa,
            iso_boyicha_yangilarini_ajrat, notanish_soni,
        )
        await msg.reply_text("⏳ Tahrirlangan fayl o'qilmoqda, biroz kuting...")
        try:
            yangilar = draft_excel_oqi(raw)
        except Exception as e:
            await msg.reply_text(t(lang, "kont_xato").format(xato=str(e)[:300]), parse_mode="Markdown")
            return
        tarix = konteyner_tarix_olish()
        yangilar = iso_boyicha_yangilarini_ajrat(yangilar, XITOY_PARSED_DIR, tarix)
        if not yangilar:
            await msg.reply_text(
                "⚠️ Faylda hech qanday to'g'ri (yoki hammasi allaqachon "
                "tasdiqlangan) qator topilmadi — qaytadan tekshiring."
            )
            return
        context.user_data["kont_yangilar"] = yangilar
        await msg.reply_text("✏️ *Yangilangan ro'yxat:*\n\n" + qisqa_xulosa(yangilar), parse_mode="Markdown")
        bio = draft_excel_yarat(yangilar)
        await msg.reply_document(
            document=bio, filename="Yangi_konteynerlar.xlsx",
            caption=(
                "Yana tahrirlash kerak bo'lsa qayta yuboring. "
                "Hammasi to'g'ri bo'lsa — pastdagi tugmani bosing."
            ),
        )
        n_nomos = notanish_soni(yangilar)
        if n_nomos:
            await msg.reply_text(
                f"⚠️ *DIQQAT: {n_nomos} ta tovar hali ham inventarda "
                f"NOTANISH* (faylda ⚠️ belgi bilan ajratilgan). Iltimos, "
                f"yuqoridagi Excel faylni oching, shu qatorlarni TO'G'IRLAB "
                f"qayta shu yerga yuboring — aks holda ular xato/noaniq nom "
                f"bilan saqlanib qoladi.",
                parse_mode="Markdown",
            )
            await msg.reply_text(
                "✅ Baribir tasdiqlaysizmi?",
                reply_markup=kont_tasdiq_ikb(lang),
            )
        else:
            await msg.reply_text(
                "✅ Tasdiqlaysizmi?",
                reply_markup=kont_tasdiq_ikb(lang),
            )
        return

    # ── Xitoy / Yuklatish fayllari uchun parse ───────────────────────────────
    try:
        ok, xato, xitoy_map, kont_rows, ombor_map, vazn_map = xitoy_ostatka_oqi(raw)
    except Exception as e:
        await msg.reply_text(f"❌ Fayl o'qishda xato: {str(e)[:300]}")
        return
    if not ok:
        await msg.reply_text(f"❌ {xato}")
        return

    if kut[0] == "xitoy_tp":
        context.user_data["xitoy_tp_data"] = {"tovarlar": xitoy_map or {}, "ombor": ombor_map or {}, "vazn": vazn_map or {}}
        context.user_data["kutilmoqda"] = ("xitoy_list", kanal)
        n = len(xitoy_map or {})
        await msg.reply_text(
            t(lang, "xitoy_tp_qabul").format(n=n) + "\n\n" + t(lang, "xitoy_fayl_kut_list"),
            parse_mode="Markdown",
        )

    elif kut[0] in ("xitoy_list", "xitoy_ostatka_fayl"):
        tp_data   = context.user_data.pop("xitoy_tp_data", {})
        tp_map    = tp_data.get("tovarlar", {})
        tp_ombor  = tp_data.get("ombor", {})
        tp_vazn   = tp_data.get("vazn", {})
        final_map   = {**tp_map,   **(xitoy_map   or {})}
        final_ombor = {**tp_ombor, **(ombor_map   or {})}
        final_vazn  = {**tp_vazn,  **(vazn_map    or {})}
        xitoy_saqlash(kanal, final_map, final_ombor)
        n_new = vazn_lookup_yangilash(final_vazn)
        if n_new:
            logger.info(f"vazn_lookup: {n_new} ta yangi tovar qo'shildi")
        context.user_data.pop("kutilmoqda", None)
        n = len(final_map)
        await msg.reply_text(
            t(lang, "xitoy_qabul").format(n=n),
            parse_mode="Markdown",
        )
        await draft_buyurtma_yubor(msg, context, kanal, lang, xitoy_ostatka=final_map)

    elif kut[0] == "xitoy_fayl":
        # DIQQAT (2026-07-08 topilgan va tuzatilgan JIDDIY XATO): bu
        # bo'lim avval BO'SH edi ("pass") — "Konteyner yuklash" →
        # "Yuklash rejasi" oqimida fayl yuborilganda, u yuqorida
        # (1536-qatorda) allaqachon o'qib bo'lingan (xitoy_ostatka_oqi)
        # bo'lsa ham, natija HECH QAYERGA yozilmasdi va foydalanuvchiga
        # HECH QANDAY javob (na xato, na tasdiq) qaytarilmasdi — bot
        # "abadiy jim" bo'lib qolganday tuyulardi. Endi natija
        # xitoy_akkum/ombor_akkum/vazn_akkum ga QO'SHILADI (bir nechta
        # fayl ketma-ket yuborilsa hammasi birlashadi) va tasdiq
        # xabari yuboriladi.
        akkum       = context.user_data.get("xitoy_akkum", {})
        ombor_akkum = context.user_data.get("ombor_akkum", {})
        vazn_akkum  = context.user_data.get("vazn_akkum", {})
        akkum.update(xitoy_map or {})
        ombor_akkum.update(ombor_map or {})
        vazn_akkum.update(vazn_map or {})
        context.user_data["xitoy_akkum"] = akkum
        context.user_data["ombor_akkum"] = ombor_akkum
        context.user_data["vazn_akkum"]  = vazn_akkum

        n = len(ombor_map or {}) or len(xitoy_map or {})
        jami = len(ombor_akkum) or len(akkum)
        await msg.reply_text(
            t(lang, "xitoy_fayl_qabul").format(n=n, jami=jami),
            parse_mode="Markdown",
        )


# ── Konteyner holati ─────────────────────────────────────────────────
def _kont_parse(kont_dir):
    yolda, keldi = [], []
    if not kont_dir.exists():
        return yolda, keldi
    for f in sorted(kont_dir.glob("*.xlsx")):
        stem = f.stem
        if stem.endswith("_D"):
            base = stem[:-2]
            parts = base.split("_", 1)
            keldi.append((parts[0], parts[1] if len(parts) > 1 else "?", f.name))
        else:
            parts = stem.split("_", 1)
            yolda.append((parts[0], parts[1] if len(parts) > 1 else "?", f.name))
    return yolda, keldi


def _main_py_ishga_tushir():
    import subprocess
    try:
        subprocess.Popen(
            ["python", str(BASE_DIR / "main.py")],
            cwd=str(BASE_DIR),
            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0,
        )
    except Exception:
        pass



def _yolda_fayllar_topish(query_key: str) -> list:
    """query_key ("sana:..." yoki "nom:...") bo'yicha YOLDA (hali kelmagan)
    konteyner fayllarini topib, sorted list qaytaradi. kg_tgl va
    _kont_list_yuborish (is_keldi=True) uchun umumiy logika."""
    if query_key.startswith("sana:"):
        sana = query_key[5:]
        return sorted(XITOY_PARSED_DIR.glob(f"*_{sana}.xlsx"))
    else:
        nom = query_key[4:].upper()
        return sorted(f for f in XITOY_PARSED_DIR.glob("*.xlsx")
                      if not f.stem.endswith("_D") and f.stem.upper().startswith(nom))


def _keldi_fayllar_topish(query_key: str) -> list:
    """query_key bo'yicha KELDI bo'lgan (allaqachon kelgan) konteyner
    fayllarini topib, sorted list qaytaradi. qt_tgl va
    _kont_list_yuborish (is_keldi=False) uchun umumiy logika."""
    if query_key.startswith("sana:"):
        sana = query_key[5:]
        return sorted(XITOY_PARSED_DIR.glob(f"*_{sana}_D.xlsx"))
    else:
        nom = query_key[4:].upper()
        return sorted(f for f in XITOY_PARSED_DIR.glob("*_D.xlsx")
                      if f.stem[:-2].upper().startswith(nom))


async def _kont_list_yuborish(msg, query_key: str, is_keldi: bool,
                              reply: bool = True, context=None):
    """
    query_key: "sana:07.06.2026" yoki "nom:CRXU"
    is_keldi=True  -> YOLDA larni korsatib KELDI tugmalari
    is_keldi=False -> KELDI larni korsatib QAYTARISH tugmalari
    """
    tag = query_key.split(":", 1)[1] if ":" in query_key else query_key

    if query_key.startswith("sana:"):
        sana = query_key[5:]
        if is_keldi:
            fayllar = _yolda_fayllar_topish(query_key)
            sarlavha = f"\U0001f6a2 *{sana}* \u2014 yo'ldagi konteynerlar:"
        else:
            fayllar = _keldi_fayllar_topish(query_key)
            sarlavha = f"\u2705 *{sana}* \u2014 kelgan konteynerlar:"
    else:
        nom = query_key[4:].upper()
        if is_keldi:
            fayllar = _yolda_fayllar_topish(query_key)
            sarlavha = f"\U0001f6a2 *{nom}* \u2014 yo'ldagi konteynerlar:"
        else:
            fayllar = _keldi_fayllar_topish(query_key)
            sarlavha = f"\u2705 *{nom}* \u2014 kelgan konteynerlar:"

    if not fayllar:
        topilmadi = f"\u274c *{tag}* bo'yicha konteyner topilmadi."
        await msg.reply_text(topilmadi, parse_mode="Markdown")
        if context is not None:
            context.user_data["screen"] = "keldi_menu"
            await kont_holat_royhat(msg, context)
        return

    # \u2500\u2500 Bir nechta konteyner bir vaqtda KELDI bo'lsa \u2014 belgilash (checkbox) rejimi \u2500\u2500
    if is_keldi and len(fayllar) > 1:
        if context is not None:
            context.user_data["kg_multi_sel"] = {"query_key": query_key, "selected": set()}
        await msg.reply_text(
            sarlavha + "\n\n_Kerakli konteynerlarni belgilang, so'ng pastdagi tugmani bosing:_",
            parse_mode="Markdown",
            reply_markup=_kg_multi_kb(fayllar, query_key, set()),
        )
        return

    # \u2500\u2500 Bir nechta konteyner bir vaqtda QAYTARISH kerak bo'lsa \u2014 checkbox rejimi \u2500\u2500
    if not is_keldi and len(fayllar) > 1:
        if context is not None:
            context.user_data["qt_multi_sel"] = {"query_key": query_key, "selected": set()}
        await msg.reply_text(
            sarlavha + "\n\n_Qaytariladigan konteynerlarni belgilang, so'ng pastdagi tugmani bosing:_",
            parse_mode="Markdown",
            reply_markup=_qt_multi_kb(fayllar, query_key, set()),
        )
        return

    buttons = []
    for f in fayllar:
        if is_keldi:
            iso    = f.stem.rsplit("_", 1)[0]
            sana_f = f.stem.rsplit("_", 1)[-1]
            lbl    = f"\U0001f6a2 {iso} ({sana_f})"
            cb     = f"kont_bir_keldi:{f.name}|{query_key}"
        else:
            stem_no_d = f.stem[:-2]
            iso    = stem_no_d.rsplit("_", 1)[0]
            sana_f = stem_no_d.rsplit("_", 1)[-1]
            lbl    = f"\u2705 {iso} ({sana_f})"
            cb     = f"kont_bir_qayt:{f.name}|{query_key}"
        buttons.append([InlineKeyboardButton(lbl, callback_data=cb)])

    await msg.reply_text(sarlavha, parse_mode="Markdown",
                         reply_markup=InlineKeyboardMarkup(buttons))


def _kg_multi_kb(fayllar, query_key: str, selected: set) -> InlineKeyboardMarkup:
    """YOLDA konteynerlar ro'yxati \u2014 checkbox (belgilash) klaviaturasi."""
    buttons = []
    for f in fayllar:
        iso    = f.stem.rsplit("_", 1)[0]
        sana_f = f.stem.rsplit("_", 1)[-1]
        mark   = "\u2705" if f.name in selected else "\u2b1c"
        lbl    = f"{mark} {iso} ({sana_f})"
        cb     = f"kg_tgl:{f.name}|{query_key}"
        buttons.append([InlineKeyboardButton(lbl, callback_data=cb)])
    n = len(selected)
    buttons.append([InlineKeyboardButton(
        f"\u2705 Tanlanganlarni KELDI qilish ({n} ta)" if n else "\u2705 Tanlanganlarni KELDI qilish",
        callback_data=f"kg_confirm_multi:{query_key}"
    )])
    buttons.append([InlineKeyboardButton("\u274c Bekor qilish", callback_data="kg_cancel_multi")])
    return InlineKeyboardMarkup(buttons)


def _qt_multi_kb(fayllar, query_key: str, selected: set) -> InlineKeyboardMarkup:
    """KELDI konteynerlar ro'yxati \u2014 QAYTARISH uchun checkbox klaviaturasi."""
    buttons = []
    for f in fayllar:
        stem_no_d = f.stem[:-2]
        iso    = stem_no_d.rsplit("_", 1)[0]
        sana_f = stem_no_d.rsplit("_", 1)[-1]
        mark   = "\u2705" if f.name in selected else "\u2b1c"
        lbl    = f"{mark} {iso} ({sana_f})"
        cb     = f"qt_tgl:{f.name}|{query_key}"
        buttons.append([InlineKeyboardButton(lbl, callback_data=cb)])
    n = len(selected)
    buttons.append([InlineKeyboardButton(
        f"\u21a9\ufe0f Tanlanganlarni qaytarish ({n} ta)" if n else "\u21a9\ufe0f Tanlanganlarni qaytarish",
        callback_data=f"qt_confirm_multi:{query_key}"
    )])
    buttons.append([InlineKeyboardButton("\u274c Bekor qilish", callback_data="qt_cancel_multi")])
    return InlineKeyboardMarkup(buttons)


async def _kg_yuborish_guruhga(msg, context, iso: str, file_id: str, caption: str):
    """KELDI bo'lgan konteyner rasmini 'Kelgan yuklar' guruhi/mavzusiga yuboradi."""
    if not config.KELGAN_YUKLAR_CHAT_ID:
        await msg.reply_text(
            "⚠️ Guruh sozlanmagan. .env faylida KELGAN_YUKLAR_CHAT_ID "
            "(va kerak bo'lsa KELGAN_YUKLAR_TOPIC_ID) qiymatini kiriting."
        )
    else:
        try:
    