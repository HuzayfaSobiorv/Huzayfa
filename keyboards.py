"""
keyboards.py — Barcha Telegram klaviatura funksiyalari
"""
from telegram import (
    InlineKeyboardButton, InlineKeyboardMarkup,
    ReplyKeyboardMarkup, KeyboardButton,
)
from texts import t


# ── Qurilish yordamchilari ───────────────────────────────────────────────────

def rkb(*rows, one_time: bool = False) -> ReplyKeyboardMarkup:
    """Reply keyboard (pastki klaviatura) yaratadi."""
    return ReplyKeyboardMarkup(
        [[KeyboardButton(lbl) for lbl in row] for row in rows],
        resize_keyboard=True,
        one_time_keyboard=one_time,
    )


def ikb(*rows) -> InlineKeyboardMarkup:
    """Inline keyboard yaratadi. Har row: [(label, callback_data), ...]"""
    return InlineKeyboardMarkup(
        [[InlineKeyboardButton(lbl, callback_data=cb) for lbl, cb in row]
         for row in rows]
    )


# ── Reply klaviaturalar ──────────────────────────────────────────────────────

def main_kb(lang: str) -> ReplyKeyboardMarkup:
    """Admin uchun to'liq menyu."""
    return rkb(
        [t(lang, "b_order"),      t(lang, "b_load")],
        [t(lang, "b_konteyner"),  t(lang, "b_karta")],
        [t(lang, "b_settings")],
    )


def konteyner_kb(lang: str) -> ReplyKeyboardMarkup:
    """Konteynerlar bilan ishlash ekrani."""
    return rkb(
        [t(lang, "b_yolda_excel")],
        [t(lang, "b_yolga_kont")],
        [t(lang, "b_keldi_belgi")],
        [t(lang, "back")],
    )


def main_kb_user(lang: str) -> ReplyKeyboardMarkup:
    """Filyal boshqaruvchisi uchun soddalashtirilgan menyu."""
    return rkb(
        [t(lang, "b_yolda_excel"), t(lang, "b_karta")],
        [t(lang, "b_settings")],
    )


def settings_kb_user(lang: str) -> ReplyKeyboardMarkup:
    """Filyal foydalanuvchisi sozlamalari."""
    return rkb(
        [t(lang, "b_lang")],
        [t(lang, "b_sorov_yuborish")],
        [t(lang, "b_boglanish")],
        [t(lang, "back")],
    )


def boglanish_ikb(lang: str, phone: str = "", username: str = "") -> InlineKeyboardMarkup | None:
    """Bog'lanish inline tugmasi — faqat Telegram username (tel: URL ishlamaydi)."""
    if username:
        lbl = t(lang, "boglanish_tg")
        url = f"https://t.me/{username.lstrip('@')}"
        return InlineKeyboardMarkup([[InlineKeyboardButton(lbl, url=url)]])
    return None


def order_kb(lang: str) -> ReplyKeyboardMarkup:
    return rkb(
        [t(lang, "ch_asosiy")],
        [t(lang, "ch_sex")],
        [t(lang, "ch_osh")],
        [t(lang, "back")],
    )


def order_channel_kb(lang: str) -> ReplyKeyboardMarkup:
    return rkb(
        [t(lang, "b_excel")],
        [t(lang, "b_tasdiq")],
        [t(lang, "back")],
    )


def load_kb(lang: str) -> ReplyKeyboardMarkup:
    return rkb(
        [t(lang, "ch_asosiy")],
        [t(lang, "ch_sex")],
        [t(lang, "ch_osh")],
        [t(lang, "back")],
    )


def load_channel_kb(lang: str) -> ReplyKeyboardMarkup:
    return rkb(
        [t(lang, "back")],
    )


def status_kb(lang: str) -> ReplyKeyboardMarkup:
    return rkb(
        [t(lang, "b_karta")],
        [t(lang, "b_yolda")],
        [t(lang, "back")],
    )


def settings_kb(lang: str, admin: bool = False) -> ReplyKeyboardMarkup:
    rows = [
        [t(lang, "b_yangilash")],
        [t(lang, "b_lang")],
        [t(lang, "b_tozala_buy")],
        [t(lang, "b_tozala_xitoy")],
        [t(lang, "b_sorovlar_royxat")],
    ]
    rows.append([t(lang, "back")])
    return rkb(*rows)


def kont_tasdiq_ikb(lang: str) -> InlineKeyboardMarkup:
    return ikb(
        [(t(lang, "kont_tasdiq_ha"), "kont:ha")],
        [(t(lang, "kont_tasdiq_yoq"), "kont:yoq")],
    )


def search_kb(lang: str) -> ReplyKeyboardMarkup:
    return rkb([t(lang, "back")])


# ── Inline klaviaturalar ─────────────────────────────────────────────────────

def til_ikb() -> InlineKeyboardMarkup:
    return ikb(
        [("Кирилл (Ўзбекча)", "lang:cyr")],
        [("Lotin (O'zbekcha)", "lang:lat")],
    )


def grafik_kat_ikb() -> InlineKeyboardMarkup:
    return ikb(
        [("🔘 Труба",        "karta_kat:truba"),  ("🔲 Профиль",       "karta_kat:profil")],
        [("📄 Лист",         "karta_kat:list"),   ("📏 Баласина",       "karta_kat:bal")],
        [("🔩 Стойка",       "karta_kat:stoyka"), ("🎯 Чашка",          "karta_kat:chas")],
        [("🍄 Қузиқорин",    "karta_kat:kuz"),    ("🔵 Шар",            "karta_kat:shar")],
        [("💿 Соққа",        "karta_kat:sokka"),  ("🪟 Ойна держатель", "karta_kat:oyna")],
        [("🔍 Umumiy qidiruv", "karta_umumiy")],
    )


def grafik_tovar_ikb(tovars: list) -> InlineKeyboardMarkup:
    rows = []
    for i, tov in enumerate(tovars[:10]):
        short = tov if len(tov) <= 45 else tov[:42] + '...'
        rows.append([(short, f"karta_tovar:{i}")])
    return InlineKeyboardMarkup(
        [[InlineKeyboardButton(lbl, callback_data=cb)] for lbl, cb in [r[0] for r in rows]]
    )


def xitoy_sorash_ikb(lang: str, kanal: str) -> InlineKeyboardMarkup:
    return ikb(
        [(t(lang, "xitoy_ha"),  f"xitoy:ha:{kanal}")],
        [(t(lang, "xitoy_yoq"), f"xitoy:yoq:{kanal}")],
    )


def xitoy_mavjud_ikb(lang: str, kanal: str) -> InlineKeyboardMarkup:
    # "Hisobsiz ber" yo'q — yangi ostatka yuklansa eski zakaz+ostatka avtomatik o'chadi,
    # shuning uchun hisobsiz berish logikasi mavjud emas.
    return ikb(
        [(t(lang, "xitoy_ishlatsin"), f"xitoy:ishlatsin:{kanal}")],
        [(t(lang, "xitoy_yangi"),     f"xitoy:yangi:{kanal}")],
    )


def xitoy_yana_ikb(lang: str, kanal: str, jami: int, tip: str = "xitoy_fayl") -> InlineKeyboardMarkup:
    # Faqat yuklatish rejasi flow uchun (xitoy_fayl) ishlatiladi
    tayyor_lbl  = (f"▶️ Hisoblashni boshlash — {jami} ta" if lang == "lat"
                   else f"▶️ ᒒисоблашни бошлаш — {jami} та")
    tayyor_data = f"xitoy:hisoblash:{kanal}"
    yana_data   = f"xitoy:yana_f:{kanal}"
    return ikb(
        [(t(lang, "xitoy_yana_btn"), yana_data)],
        [(tayyor_lbl,                 tayyor_data)],
    )


def tozala_kanal_ikb(lang: str, tip: str) -> InlineKeyboardMarkup:
    bekor = "❌ Bekor" if lang == "lat" else "❌ Бекор"
    return ikb(
        [("🏢 Asosiy", f"tozala_{tip}:asosiy"),
         ("🏭 Tsex",   f"tozala_{tip}:sex"),
         ("🇰🇬 O'sh",  f"tozala_{tip}:osh")],
        [(bekor, "tozala_no")],
    )


def tozala_tasdiq_ikb(lang: str, tip: str, kanal: str) -> InlineKeyboardMarkup:
    return ikb(
        [(t(lang, "tozala_ha"),  f"tozala_{tip}_ok:{kanal}")],
        [("❌ Bekor" if lang == "lat" else "❌ Бекор", "tozala_no")],
    )


def zakaz_tasdiq_ikb(lang: str, kanal: str) -> InlineKeyboardMarkup:
    return ikb(
        [(t(lang, "zakaz_tasdiq_btn"), f"zakaz_ok:{kanal}")],
        [(t(lang, "zakaz_bekor_btn"), "zakaz_no")],
    )
