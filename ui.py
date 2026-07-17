"""
ui.py — Foydalanuvchi interfeysi: animatsiya, ekranlar, grafik, viewlar
"""
import asyncio, logging, math, re
from contextlib import asynccontextmanager
from datetime import datetime
from io import BytesIO

import pandas as pd
from config import BASE_DIR, DATA_FILE, BACK_MAP, CH_KEY, _MPL_LOCK, get_inv, get_kont
from texts import t
from keyboards import (
    main_kb, main_kb_user, order_kb, order_channel_kb, load_kb, load_channel_kb,
    status_kb, settings_kb, settings_kb_user, search_kb, konteyner_kb,
    grafik_kat_ikb, grafik_tovar_ikb,
    xitoy_sorash_ikb, xitoy_mavjud_ikb, boglanish_ikb,
)
from services import (
    inventar_olish, kamomat_olish, kritiklar_text,
    buyurtma_yuklash, xitoy_yuklash, asosiy_styled_excel_yarat,
    draft_excel_yarat, grafik_qidirish, qidiruv_text,
    xlsx_mavjud, rasm_pending_iso_royxati,
)
from kamomat_engine import kamomat_stats_v2, kamomat_excel_v2
from yolda_excel import yolda_excel

# Lokal aliaslar — config funksiyalari
_get_inv  = get_inv
_get_kont = get_kont

logger = logging.getLogger(__name__)


async def _typing_loop(chat_id: int, bot, stop: asyncio.Event, action: str = "upload_document"):
    """Telegram header'ida typing/upload indikatorini yangilaydi (4 sek)."""
    while not stop.is_set():
        try:
            await bot.send_chat_action(chat_id=chat_id, action=action)
        except Exception:
            pass
        try:
            await asyncio.wait_for(asyncio.shield(stop.wait()), timeout=4)
        except (asyncio.TimeoutError, asyncio.CancelledError):
            pass


async def _dots_animation(msg, text: str, stop: asyncio.Event):
    """
    Xabarda animatsion nuqtalar: ⏳ text. → ⏳ text.. → ⏳ text... → ⏳ text.
    stop.set() da loop tugaydi, xabar o'chadi.
    """
    anim_msg = None
    try:
        anim_msg = await msg.reply_text(f"⏳ {text}.")
        dot = 1
        while not stop.is_set():
            await asyncio.sleep(0.3)
            if stop.is_set():
                break
            dot = dot % 3 + 1          # 1→2→3→1 (hech qachon 0 bo'lmaydi)
            try:
                await anim_msg.edit_text(f"⏳ {text}{'.' * dot}")
            except Exception:
                pass
    except asyncio.CancelledError:
        pass
    finally:
        if anim_msg:
            try:
                await anim_msg.delete()
            except Exception:
                pass


@asynccontextmanager
async def yuklash_animatsiya(msg, context, text_cyr: str = "", text_lat: str = "",
                              action: str = "upload_document"):
    """
    Matn berilsa → xabarda animatsion nuqtalar (⏳ text...).
    Matn berilmasa → faqat Telegram header typing indikatori.
    stop.set() dan keyin tasklar o'zlari tugaydi (cancel yo'q) — xabar ishonchli o'chadi.
    """
    stop = asyncio.Event()
    tasks = []

    tasks.append(asyncio.create_task(
        _typing_loop(msg.chat_id, context.bot, stop, action)
    ))

    lang = context.user_data.get("lang", "cyr")
    anim_text = text_cyr if lang == "cyr" else text_lat
    if anim_text:
        tasks.append(asyncio.create_task(_dots_animation(msg, anim_text, stop)))

    try:
        yield
    finally:
        stop.set()
        # Cancel yo'q — tasklar stop.is_set() ko'rib o'zlari chiqadi (≤0.3 sek)
        # _dots_animation finally bloki xabarni o'chiradi
        for t_ in tasks:
            try:
                await asyncio.wait_for(t_, timeout=1.5)
            except (asyncio.TimeoutError, Exception):
                t_.cancel()
                try:
                    await t_
                except (asyncio.CancelledError, Exception):
                    pass



def _is_admin(context) -> bool:
    # 2026-07-16: config.ADMIN_IDS EMAS — admin_idlari() ishlatiladi, chunki
    # endi admin ro'yxati dinamik (/addadmin, /removeadmin bilan boshqariladi,
    # .env kerak emas). Qarang: services.py::admin_idlari.
    from services import admin_idlari
    uid = context.user_data.get("user_id") if context else None
    return uid is not None and uid in admin_idlari()


def build_screen(screen: str, lang: str, context) -> tuple:
    """(matn, klaviatura) qaytaradi."""
    kanal = context.user_data.get("kanal", "asosiy")
    admin = _is_admin(context)

    if screen == "main":
        if admin:
            return t(lang, "main_title"), main_kb(lang)
        return t(lang, "user_main_title"), main_kb_user(lang)
    if screen == "order":
        return t(lang, "order_title"), order_kb(lang)
    if screen == "order_channel":
        ch = t(lang, CH_KEY[kanal])
        return f"{ch}\n\n{t(lang, 'order_ch_title')}", order_channel_kb(lang)
    if screen == "load":
        return t(lang, "load_title"), load_kb(lang)
    if screen == "load_channel":
        ch = t(lang, CH_KEY[kanal])
        return t(lang, "load_ch_title").format(ch=ch), load_channel_kb(lang)
    if screen == "status":
        return t(lang, "status_title"), status_kb(lang)
    if screen == "settings":
        if admin:
            return t(lang, "settings_title"), settings_kb(lang, admin=True)
        return t(lang, "settings_title_user"), settings_kb_user(lang)
    if screen in ("search", "search_kat"):
        return t(lang, "search_title"), search_kb(lang)
    if screen == "konteyner":
        return t(lang, "konteyner_title"), konteyner_kb(lang)
    if screen == "keldi_ekran":
        from telegram import ReplyKeyboardMarkup
        orqaga_kb = ReplyKeyboardMarkup([[t(lang, "back")]], resize_keyboard=True)
        return t(lang, "konteyner_title"), orqaga_kb
    return t(lang, "user_main_title") if not admin else t(lang, "main_title"), \
           main_kb_user(lang) if not admin else main_kb(lang)


async def aktiv_inline_tozala(context, bot):
    """2026-07-17 (Huzayfa: userlar eski ekrandan qolib ketgan inline
    tugmalarni bosib chalkashib ketyapti — masalan qidiruvda Труба tugmasi
    hali ko'rinib turibdi, keyin Sozlamalarga o'tib o'sha eski tugmani
    bosadi va bot buni joriy ekranga tegishli deb noto'g'ri qayta ishlaydi).

    Har safar foydalanuvchi YANGI ekranga o'tganda shu funksiya chaqiriladi:
    oldingi ekrandan qolgan inline (bosiladigan) tugmalarni FIZIK o'chiradi,
    shunda eski xabar oddiy matnga aylanadi va boshqa bosib bo'lmaydi.
    """
    aktiv = context.user_data.pop("aktiv_inline", None)
    if not aktiv:
        return
    chat_id, message_id = aktiv
    try:
        await bot.edit_message_reply_markup(chat_id=chat_id, message_id=message_id, reply_markup=None)
    except Exception:
        pass   # xabar allaqachon o'zgargan / o'chirilgan / juda eski — muammo emas


def aktiv_inline_belgila(context, sent_msg):
    """Yuborilgan xabarda inline tugma bo'lsa, uni 'joriy faol ekran' deb
    belgilaydi — keyingi ekran o'tishida shu xabar tozalanadi (yuqoridagi
    aktiv_inline_tozala orqali)."""
    try:
        if sent_msg is not None and sent_msg.reply_markup is not None:
            context.user_data["aktiv_inline"] = (sent_msg.chat_id, sent_msg.message_id)
    except Exception:
        pass


async def go_screen(msg, context, screen: str, kanal: str | None = None):
    """Yangi ekranga o'tadi va reply keyboard ni yangilaydi."""
    await aktiv_inline_tozala(context, msg.get_bot())
    lang = context.user_data.get("lang", "cyr")
    if kanal:
        context.user_data["kanal"] = kanal
    context.user_data["screen"] = screen
    text, kb = build_screen(screen, lang, context)
    await msg.reply_text(text, reply_markup=kb, parse_mode="Markdown")


def get_action(lang: str, screen: str, text: str):
    """
    Ekran + bosilgan matn → amal kalit so'zi.
    Tuple: (keyingi_ekran, kanal) — kanal ekraniga o'tish.
    String: amal nomi.
    None:  noma'lum.
    """
    MAP = {
        "main": {
            t(lang, "b_order"):       "order",
            t(lang, "b_load"):        "load",
            t(lang, "b_karta"):       "karta",
            t(lang, "b_yolda"):       "yolda",
            t(lang, "b_konteyner"):   "konteyner",
            t(lang, "b_settings"):    "settings",
            # User (filyal) tugmalari
            t(lang, "b_yolda_excel"): "yolda_excel",
        },
        "konteyner": {
            t(lang, "b_yolda_excel"): "yolda_excel",
            t(lang, "b_yolga_kont"):  "yolga_kont",
            t(lang, "b_keldi_belgi"): "keldi_belgi",
        },
        "order": {
            t(lang, "ch_asosiy"): ("order_channel", "asosiy"),
            t(lang, "ch_sex"):    ("order_channel", "sex"),
            t(lang, "ch_osh"):    ("order_channel", "osh"),
        },
        "order_channel": {
            t(lang, "b_excel"):   "excel",
            t(lang, "b_tasdiq"):  "tasdiq",
        },
        "load": {
            t(lang, "ch_asosiy"): ("load_channel", "asosiy"),
            t(lang, "ch_sex"):    ("load_channel", "sex"),
            t(lang, "ch_osh"):    ("load_channel", "osh"),
        },
        "load_channel": {
            t(lang, "b_hisoblash"): "hisoblash",
        },
        "status": {
            t(lang, "b_karta"): "karta",
            t(lang, "b_yolda"): "yolda",
        },
        "settings": {
            t(lang, "b_yangilash"):     "yangilash",
            t(lang, "b_lang"):          "lang_pick",
            t(lang, "b_tozala_buy"):    "tozala_buy",
            t(lang, "b_tozala_xitoy"):  "tozala_xitoy",
            t(lang, "b_yolga_kont"):    "yolga_kont",
            t(lang, "b_sorovlar_royxat"): "sorovlar_royxat",
            # User settings
            t(lang, "b_boglanish"):     "boglanish",
            t(lang, "b_sorov_yuborish"): "sorov_yuborish",
        },
    }
    return MAP.get(screen, {}).get(text)


async def grafik_ko_rsatish(msg, tovar: str, kanal: str, kat: str = "truba"):
    """Tovar kartasi: matn zudlik bilan, grafik keyin alohida."""
    from kamomat_engine import grafik_chiz, zanjir_sim
    loop = asyncio.get_event_loop()

    # ── 1. Ma'lumotni RAM cache dan olamiz ───────────────────────────────────
    def _norm(s: str) -> str:
        """Товар nomini normalize qiladi: 'Лист- 0,8' → 'Лист-0,8'"""
        import re
        return re.sub(r'-\s+', '-', s.strip())

    def _data_olish():
        inv = _get_inv()
        tv     = str(tovar).strip()
        tv_n   = _norm(tv)
        inv_n  = inv["Товар"].astype(str).apply(_norm)
        row    = inv[inv_n == tv_n]
        if row.empty:
            return None
        r0      = row.iloc[0]
        qoldiq  = float(r0.get("Қолдиқ", 0) or 0)
        min_z   = float(r0.get("Мин_Захира", 0) or 0)
        yolda_j = float(r0.get("Йўлда_Жами", 0) or 0)
        holat   = str(r0.get("Холат", ""))
        kont_list, kont_rows = [], []
        # 2026-07-16 (Huzayfa): rasm allaqachon guruhga yuborilgan konteyner
        # DISPLAY'da (kont_rows — "Yo'ldagi konteynerlar" matnida) ko'rinmasin
        # — lekin kont_list (buyurtma hisob-kitobi, zanjir_sim) ga TA'SIR
        # qilmaydi, xuddi Yo'lda Excel'dagi bilan bir xil mantiq.
        rasm_yuborilganlar = rasm_pending_iso_royxati()
        try:
            kont   = _get_kont()
            kont   = kont[kont["Холат"].astype(str).str.strip() != "КЕЛДИ ✅"].copy()
            kont_n = kont["Товар"].astype(str).apply(_norm)
            kont_f = kont[kont_n == tv_n].copy()
            if "Кун_Қолди" in kont_f.columns:
                kont_f = kont_f.sort_values("Кун_Қолди")
            for _, r in kont_f.iterrows():
                kq    = float(r.get("Кун_Қолди", 0) or 0)
                mq    = float(r.get("Миқдор", 0) or 0)
                kech  = float(r.get("Кечикиш_Кун", 0) or 0)
                if mq <= 0:
                    continue
                kont_list.append((kq, mq))
                kont_nomi = str(r.get("Контейнер", "—"))
                if kont_nomi in rasm_yuborilganlar:
                    continue   # hisobga kiradi, lekin ko'rsatilmaydi
                kont_rows.append({
                    "nom":    kont_nomi,
                    "miqdor": int(mq),
                    "yukl":   str(r.get("Юкланган_Сана", "—")),
                    "kelish": str(r.get("Келиш_Санаси", "—")),
                    "kun":    int(kq),
                    "kechik": int(kech),
                })
        except Exception:
            pass
        sim = zanjir_sim(qoldiq, min_z, kont_list)
        return holat, qoldiq, min_z, yolda_j, sim, kont_list, kont_rows

    try:
        result = await loop.run_in_executor(None, _data_olish)
    except Exception as e:
        await msg.reply_text(f"❌ Xato: {e}")
        return

    if result is None:
        await msg.reply_text(f"❌ Topilmadi: {tovar}")
        return

    holat, qoldiq, min_z, yolda_j, sim, kont_list, kont_rows = result

    # 2026-07-16 (Huzayfa): min-zaxira 0 bo'lgan ("meyor yo'q", buyurtmada
    # kuzatilmaydigan) tovar qidiruvda ICHKI holat (min-zaxira yo'q) sifatida
    # emas, mijozbop "aksiya mahsulot sifatida sotuvdan chiqarilgan" deb
    # ko'rsatiladi — boshqa hech qanday hisob-kitob/grafik chiqmaydi.
    if min_z <= 0:
        await msg.reply_text(
            f"🏷 *{tovar}*\n\n"
            "Bu mahsulot aksiya mahsulot sifatida sotuvdan chiqarilgan.",
            parse_mode="Markdown",
        )
        return

    uzilish = sim.get("uzilish_kun")
    taklif  = sim.get("taklif", 0)


    # ── Xavf matni ────────────────────────────────────────────
    if uzilish is None:
        xavf_txt = "✅ Uzilish yo'q"
    elif uzilish <= 0:
        xavf_txt = "🔴 HOZIR uzilish!"
    elif uzilish <= 30:
        xavf_txt = f"🔴 Uzilish: *{uzilish}* kun ichida"
    elif uzilish <= 60:
        xavf_txt = f"🟡 Uzilish: *{uzilish}* kun ichida"
    else:
        xavf_txt = f"🟢 Uzilish: *{uzilish}* kun ichida"

    # ── Konteyner qatorlari ───────────────────────────────────
    if kont_rows:
        kont_lines = ["🚢 *Yo'ldagi konteynerlar:*"]
        for kr in kont_rows:
            if kr.get('kechik', 0) > 0:
                kun_txt = f"Kechikdi: {kr['kechik']} kun ⚠️"
            elif kr['kun'] > 0:
                kun_txt = f"{kr['kun']} kun qoldi 🕐"
            else:
                kun_txt = "Bugun keladi 📬"
            kont_lines.append(
                f"📦 *{kr['nom']}* — {kr['miqdor']} dona\n"
                f"   Yuklangan: {kr['yukl']}  |  Kelish: {kr['kelish']}  |  {kun_txt}"
            )
        kont_txt = "\n\n".join(kont_lines)
    else:
        kont_txt = "🚢 Yo'lda konteyner yo'q"

    # ── 2. Matn kartasi — ZUDLIK BILAN yuboramiz ─────────────────────────────
    text_card = (
        f"📊 *{tovar}*\n\n"
        f"{holat}\n"
        f"Qoldiq: *{int(qoldiq):,}*\n"
        f"Yo'lda jami: *{int(yolda_j):,}*\n\n"
        f"▬▬▬▬▬▬▬▬▬▬▬▬\n\n"
        f"{kont_txt}\n\n"
        f"▬▬▬▬▬▬▬▬▬▬▬▬\n"
        f"{xavf_txt}"
    )
    await msg.reply_text(text_card, parse_mode="Markdown")

    # ── 3. Grafik — faqat truba, profil, list uchun ─────────────────────────
    if kat in ("truba", "profil", "list"):
        def _render_grafik():
                return grafik_chiz(tovar, qoldiq, min_z, kont_list)

        try:
            bio = await asyncio.wait_for(
                loop.run_in_executor(None, _render_grafik), timeout=30
            )
            if bio is not None:
                bio.seek(0)
                await msg.reply_photo(photo=bio, caption=f"📈 {tovar}")
        except asyncio.TimeoutError:
            logging.warning("Grafik: 30 sek timeout")
        except Exception as _g_err:
            logging.warning(f"Grafik xato: {_g_err}")



async def kamomat_ko_rish(msg, context, kanal: str, lang: str):
    ch = t(lang, CH_KEY[kanal])
    if not xlsx_mavjud():
        await msg.reply_text(t(lang,"data_yoq"), parse_mode="Markdown")
        return

    async with yuklash_animatsiya(
        msg, context,
        text_cyr="Камомат хисоблaнмоқда",
        text_lat="Kamomat hisoblanmoqda",
    ):
        # Yangi engine: KRITIK + PAST, zanjir simulyatsiyasi
        stats = kamomat_stats_v2(DATA_FILE, kanal, buyurtma_yuklash)
        if stats["n"] == 0:
            await msg.reply_text(t(lang,"kamomat_yoq"))
            return

        await msg.reply_text(
            t(lang,"kamomat_title").format(ch=ch) + "\n\n" +
            t(lang,"kamomat_line").format(**stats),
            parse_mode="Markdown",
        )

        # Rangli, tartibli Excel
        bio = kamomat_excel_v2(DATA_FILE, kanal, lang, buyurtma_yuklash)
        if bio:
            await msg.reply_document(document=bio, filename=f"Kamomat_{kanal}.xlsx")


async def draft_buyurtma_yubor(msg, context, kanal: str, lang: str,
                               xitoy_ostatka: dict | None = None):
    ch  = t(lang, CH_KEY[kanal])
    bio = None
    async with yuklash_animatsiya(
        msg, context,
        text_cyr="Буюртма Excel тайёрланмоқда",
        text_lat="Buyurtma Excel tayyorlanmoqda",
    ):
        loop = asyncio.get_event_loop()
        try:
            bio = await loop.run_in_executor(
                None, draft_excel_yarat, kanal, xitoy_ostatka
            )
        except FileNotFoundError as e:
            err = ("⚠️ Fayl topilmadi. Avval 'Malumotlarni yangilash' tugmasini bosing."
                   if lang == "lat" else
                   "⚠️ Файл топилмади. Аввал 'Маълумотларни янгилаш' тугмасини босинг.")
            logger.error(f"draft_buyurtma_yubor FileNotFoundError: {e}")
            await msg.reply_text(err)
            return
        except Exception as e:
            logger.error(f"draft_buyurtma_yubor xato: {e}", exc_info=True)
            err_msg = ("❌ Excel yaratishda xato. Qayta urinib ko'ring."
                       if lang == "lat" else
                       "❌ Excel яратишда хато. Қайта уриниб кўринг.")
            await msg.reply_text(err_msg)
            return

    if bio is None:
        await msg.reply_text(t(lang, "buyurtma_yoq"), parse_mode="Markdown")
        return

    await msg.reply_document(
        document=bio,
        filename=f"Buyurtma_taklif_{kanal}.xlsx",
        caption=t(lang,"excel_caption").format(ch=ch),
        parse_mode="Markdown",
    )


async def yolda_ko_rish(msg, context, lang: str):
    if not xlsx_mavjud():
        await msg.reply_text(t(lang,"data_yoq"), parse_mode="Markdown")
        return

    async with yuklash_animatsiya(
        msg, context, action="typing",
        text_cyr="Контейнерлар тайёрланмоқда",
        text_lat="Konteynerlar tayyorlanmoqda",
    ):
        stats: dict = {}
        # 2026-07-16: rasm allaqachon guruhga yuborilgan konteynerlar bu
        # ro'yxatda ko'rinmasin (qarang: services.py::rasm_pending_iso_royxati)
        chiqarib = rasm_pending_iso_royxati()
        bio = yolda_excel(DATA_FILE, stats=stats, chiqarib_tashlash=chiqarib)
        if bio is None:
            await msg.reply_text(t(lang, "yolda_yoq"))
            return
        # 2026-07-14: xabar endi haqiqiy xulosa beradi (ilgari yulduzchali
        # xom matn chiqardi — parse_mode berilmagan edi) va tartib to'g'risi
        # yoziladi (kechikkanlar avval, keyin kelishiga oz qolganlar).
        # 2026-07-17 (Huzayfa: qisqa va sodda bo'lsin — faqat nechta
        # konteyner yo'lda ekani ko'rinsin): n allaqachon rasm-yuborilgan
        # (guruhga tashlangan, lekin hali qo'lda KELDI qilinmagan)
        # konteynerlarni chiqarib tashlagan holda hisoblanadi (yuqoridagi
        # `chiqarib` orqali yolda_excel()ga uzatiladi).
        n = stats.get("n", 0)
        # 2026-07-17 (Huzayfa: userlarga eslatma kerak — yangi konteyner
        # yuklanganda ular allaqachon avtomatik xabar olishadi, buni shu
        # yerda ham ko'rsatib qo'yamiz, video-o'qitishda alohida aytishga
        # hojat qolmasin).
        if lang == "lat":
            caption = (
                f"Yo'lda - 🚛{n} ta\n\n"
                "Yangi konteyner yuklansa — bot sizga alohida xabar beradi."
            )
        else:
            caption = (
                f"Йўлда - 🚛{n} та\n\n"
                "Янги контейнер юкланса — бот сизга алоҳида хабар беради."
            )
        await msg.reply_document(
            document=bio,
            filename="Yolda.xlsx",
            caption=caption,
        )
