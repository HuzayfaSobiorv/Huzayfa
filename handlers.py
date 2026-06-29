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

from config import BOT_HOLAT_DIR, CH_KEY, BACK_MAP, BASE_DIR, KONT_DIR, ADMIN_IDS, SUPER_ADMIN_ID, SUPPORT_PHONE, SUPPORT_USERNAME
from config import xlsx_refresh as _xlsx_refresh
from texts import t
from keyboards import (
    main_kb, order_kb, order_channel_kb, load_kb, load_channel_kb,
    settings_kb, search_kb, til_ikb,
    xitoy_sorash_ikb, xitoy_mavjud_ikb, xitoy_yana_ikb,
    tozala_kanal_ikb, tozala_tasdiq_ikb, zakaz_tasdiq_ikb,
    grafik_kat_ikb, kont_tasdiq_ikb, boglanish_ikb,
)
from services import (
    buyurtma_yuklash, buyurtma_saqlash, buyurtma_tozala,
    pending_saqlash, pending_yuklash, pending_tozala,
    draft_saqlash, draft_yuklash,
    xitoy_yuklash, xitoy_saqlash, buyurtma_tekshir,
    kamomat_olish, zakaz_preview_text, kamomat_stats,
    grafik_qidirish, vazn_lookup_yangilash,
    kirish_ruxsati, whitelist_qosh, whitelist_ochir, whitelist_yuklash,
)
from parsers import xitoy_ostatka_oqi
from keyboards import grafik_tovar_ikb
from ui import (
    build_screen, go_screen, get_action,
    yuklash_animatsiya, grafik_ko_rsatish,
    kamomat_ko_rish, draft_buyurtma_yubor, yolda_ko_rish,
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
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            await query.message.reply_text(t(lang, "kont_tasdiq_yoq"))
        elif decision == "ha":
            yangilar = context.user_data.pop("kont_yangilar", [])
            try:
                await query.edit_message_reply_markup(reply_markup=None)
            except Exception:
                pass
            if not yangilar:
                await query.message.reply_text(t(lang, "kont_barchasi_bor"))
                return
            from konteyner_qosh import konteyner_xlsx_yarat
            n_ok = 0
            for kont in yangilar:
                try:
                    konteyner_xlsx_yarat(kont, KONT_DIR)
                    n_ok += 1
                except Exception as e:
                    logger.error(f"konteyner xlsx xato {kont['iso']}: {e}")
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
    if text == t(lang, "back"):
        parent = BACK_MAP.get(screen, "main")
        context.user_data.pop("kutilmoqda", None)
        context.user_data.pop("grafik_natijalar", None)
        await go_screen(msg, context, parent)
        # search_kat dan search ga qaytganda kategoriya tanlash qayta ko'rsatiladi
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
        return

    # Oddiy navigatsiya
    if action in ("order","load","settings","search","main"):
        await go_screen(msg, context, action)
        return

    # Buyurtma amallari
    kanal = context.user_data.get("kanal","asosiy")

    if action == "kamomat":
        await kamomat_ko_rish(msg, context, kanal, lang)

    elif action == "excel":
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
                xlsx_path = main_with_data(kanal, ombor_akkum)
            except Exception as e:
                await msg.reply_text(
                    f"❌ Yuklatish rejasi yaratishda xato:\n{str(e)[:300]}"
                )
                return
        context.user_data.pop("xitoy_akkum", None)
        context.user_data.pop("ombor_akkum", None)
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
        await yolda_ko_rish(msg, context, lang)

    elif action == "yangilash":
        loop = asyncio.get_event_loop()
        async with yuklash_animatsiya(
            msg, context,
            text_cyr="Маълумотлар янгиланмоқда",
            text_lat="Ma'lumotlar yangilanmoqda",
        ):
            try:
                await loop.run_in_executor(None, lambda: _xlsx_refresh(force=True))
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
    lang  = context.user_data.get("lang", "cyr")
    msg   = update.message
    kut   = context.user_data.get("kutilmoqda")

    if not (isinstance(kut, tuple) and kut[0] in ("xitoy_fayl", "xitoy_ostatka_fayl", "xitoy_tp", "xitoy_list", "buyurtma_tasdiq", "kont_tp", "kont_list")):
        await msg.reply_text("Fayl kutilmayapti. Avval kanal tanlang.")
        return

    kanal = kut[1]
    doc   = msg.document
    if doc is None:
        await msg.reply_text("Hujjat topilmadi.")
        return

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
            preview = f"✅ {len(items)} ta tovar buyurtmasi qabul qilindi."
        await msg.reply_text(preview, parse_mode="Markdown", reply_markup=zakaz_tasdiq_ikb(lang, kanal))
        return

    # ── Xitoy ostatka fayli (2-fayl: xitoy_tp → xitoy_list) ─────────────────
    ok, xato, xitoy_map, kont_rows, ombor_map, vazn_map = xitoy_ostatka_oqi(raw)
    if not ok:
        await msg.reply_text(f"Xato: {xato}")
        return

    if kut[0] == "xitoy_tp":
        # 1-fayl: Труба/Профиль — vaqtincha saqlash, Лист so'rash
        context.user_data["xitoy_tp_data"] = dict(xitoy_map or {})
        context.user_data["xitoy_tp_vazn"] = dict(vazn_map or {})
        context.user_data["kutilmoqda"] = ("xitoy_list", kanal)
        n = len(xitoy_map or {})
        await msg.reply_text(
            t(lang, "xitoy_tp_qabul").format(n=n) + "\n\n" + t(lang, "xitoy_fayl_kut_list"),
            parse_mode="Markdown",
        )

    elif kut[0] in ("xitoy_list", "xitoy_ostatka_fayl"):
        # 2-fayl: List Xitoy fayli yoki ostatka fayili
        tp_data   = context.user_data.pop("xitoy_tp_data", {})
        tp_map    = tp_data.get("tovarlar", {})
        tp_ombor  = tp_data.get("ombor", {})
        tp_vazn   = tp_data.get("vazn", {})
        # Ikki faylni birlashtirish
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
        # Avtomatik draft Excel yaratish
        await draft_buyurtma_yubor(msg, context, kanal, lang, xitoy_ostatka=final_map)

    elif kut[0] == "xitoy_fayl":
        pass  # fayl_keldi da ishlanadi
