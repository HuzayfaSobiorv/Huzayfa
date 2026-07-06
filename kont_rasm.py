"""
kont_rasm.py
============
Bitta konteyner "KELDI" bo'lganda uning tovar ro'yxatini
"Kelgan yuklar" guruhiga yuborish uchun PNG rasm qilib chizadi.
Uslub (ranglar, ustunlar) yolda_excel.py bilan bir xil — faqat
Excel emas, to'g'ridan-to'g'ri rasm (Pillow) qilib chiqadi.
"""

import io
from datetime import date
from pathlib import Path

import matplotlib
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

from config import DATA_FILE
from vazn_hisobla import tovar_vazni
from yolda_excel import (
    CAT_ORDER, ROW_CLR_DARK, ROW_CLR_LIGHT,
    CLR_YOLDA_BG, CLR_YOLDA_SUB, CLR_TOTAL_BG, CLR_COL_HDR_TEXT,
)

# ── Kelgan (arrived) sarlavha rangi — "Yo'ldagi konteynerlar" Excel bilan
# BIR XIL ko'k (rang uyg'unligi uchun, alohida yashil emas) ─────────────────
CLR_ARRIVED_BG   = CLR_YOLDA_BG
CLR_ARRIVED_TEXT = "FFFFFF"

FONT_DIR = Path(matplotlib.get_data_path()) / "fonts" / "ttf"

# DIQQAT: asosiy matn shrifti (DejaVuSans, matplotlib bilan birga keladi,
# har doim mavjud) 🚛/✅/⚖️ kabi RANGLI emoji belgilarni UMUMAN chiza
# olmaydi (bo'sh joy chiqadi) — bular alohida rangli shrift talab qiladi.
# Shu sababli loyihaga birga ("fonts/NotoColorEmoji.ttf") shu shrift
# ULANGAN — bot qaysi serverda ishlamasin, birga ko'chib boradi. Agar biror
# sabab bilan bu fayl topilmasa yoki Pillow/FreeType uni chiza olmasa —
# _emoji_icon() xatoga chidamli (None qaytaradi), chaqiruvchi tomon oddiy
# matn/belgi bilan ALMASHTIRADI — bot HECH QACHON shu sabab bilan yiqilmaydi.
EMOJI_FONT_PATH = Path(__file__).resolve().parent / "fonts" / "NotoColorEmoji.ttf"
_emoji_kesh: dict = {}


def _font(size: int, bold: bool = False) -> "ImageFont.FreeTypeFont":
    name = "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf"
    try:
        return ImageFont.truetype(str(FONT_DIR / name), size)
    except Exception:
        return ImageFont.load_default()


def _emoji_icon(ch: str, target_h: int) -> "Image.Image | None":
    """
    Rangli emoji belgini (masalan '🚛', '✅', '⚖️') alohida, shaffof (RGBA)
    kichik rasm sifatida qaytaradi — balandligi aniq `target_h` piksel,
    eni asl nisbatga mos ravishda hisoblanadi. Topilmasa/chizib bo'lmasa —
    None (chaqiruvchi oddiy matnga tushishi kerak).
    """
    kalit = (ch, target_h)
    if kalit in _emoji_kesh:
        return _emoji_kesh[kalit]
    icon = None
    try:
        if EMOJI_FONT_PATH.exists():
            font = ImageFont.truetype(str(EMOJI_FONT_PATH), 109)
            tmp = Image.new("RGBA", (160, 160), (255, 255, 255, 0))
            d = ImageDraw.Draw(tmp)
            d.text((12, 12), ch, font=font, embedded_color=True)
            bbox = tmp.getbbox()
            if bbox:
                cropped = tmp.crop(bbox)
                w, h = cropped.size
                new_w = max(1, round(w * target_h / h))
                icon = cropped.resize((new_w, target_h), Image.LANCZOS)
    except Exception:
        icon = None
    _emoji_kesh[kalit] = icon
    return icon


def _wrap(draw, text: str, font, max_width: int) -> list[str]:
    words = str(text).split()
    if not words:
        return [""]
    lines, cur = [], words[0]
    for w in words[1:]:
        test = f"{cur} {w}"
        if draw.textlength(test, font=font) <= max_width:
            cur = test
        else:
            lines.append(cur)
            cur = w
    lines.append(cur)
    return lines


def generate_kelgan_rasm(iso: str) -> io.BytesIO | None:
    """
    iso — konteyner raqami (masalan CRXU1561318).
    DATA_FILE dagi "Контейнерлар" varag'idan shu konteynerning
    tovar ro'yxatini o'qib, PNG rasm (BytesIO) qaytaradi.
    Ma'lumot topilmasa — None.
    """
    try:
        df = pd.read_excel(DATA_FILE, sheet_name="Контейнерлар")
    except Exception:
        return None

    grp = df[df["Контейнер"] == iso].copy()
    if grp.empty:
        return None

    first     = grp.iloc[0]
    yukl_sana = str(first.get("Юкланган_Сана", ""))[:10]
    turi      = str(first.get("Тури", ""))

    # DIQQAT: bu funksiya ADMIN "KELDI" tugmasini bosgan ANIQ shu daqiqada
    # chaqiriladi (handlers.py) — shuning uchun "Келиш_Санаси" ustunidagi
    # (Юкланган+55 kun asosida hisoblangan, TAXMINIY) sanani EMAS, balki
    # HOZIRGI, HAQIQIY kunni ko'rsatamiz. "55 kunlik" taxmin jadval va
    # boshqa hisoblarda to'liq foydali va o'zgarishsiz qoladi — FAQAT shu
    # rasmda, "kelgan" deb e'lon qilingan real kun ko'rsatiladi (kechikish/
    # tezlashish — 20, 55 yoki 65 kun bo'lishidan qat'iy nazar).
    keldi_sana = date.today().strftime("%d.%m.%Y")

    if "Категория" in grp.columns:
        grp["_ord"] = grp["Категория"].apply(lambda c: CAT_ORDER.get(str(c).strip(), 99))
        grp = grp.sort_values(["_ord", "Товар"]).drop(columns=["_ord"])

    # ── Umumiy tonna — "Вазн_кг" ustuni bo'yicha (bir marta hisoblangan,
    # konteyner qo'shilganda saqlangan), bo'lmasa tovar nomidan fallback ──
    has_vazn_col = "Вазн_кг" in grp.columns
    jami_kg = 0.0
    for _, r in grp.iterrows():
        vazn_kg = r.get("Вазн_кг") if has_vazn_col else None
        if pd.notna(vazn_kg) and vazn_kg not in (None, ""):
            jami_kg += float(vazn_kg)
        else:
            vazn = tovar_vazni(str(r.get("Товар", "")))
            miq_t = r.get("Миқдор", 0)
            if vazn and pd.notna(miq_t):
                jami_kg += vazn * miq_t
    jami_tonna = round(jami_kg / 1000, 2)

    # ── Ustun kengliklari (px) ───────────────────────────────────────────────
    PAD        = 14
    COL_NAME_W = 760
    COL_MIQ_W  = 170
    IMG_W      = PAD * 2 + COL_NAME_W + COL_MIQ_W

    HDR_H  = 64
    SUB_H  = 48
    ROW_H0 = 52   # 1-qatorli minimal balandlik (tovar nomi kattaroq bo'lgani uchun oshirildi)
    LINE_H = 34   # har bir qo'shimcha satr uchun
    TOT_H  = 46

    font_hdr  = _font(20, bold=True)
    font_sub  = _font(17, bold=True)
    font_name = _font(19)             # tovar nomi — katagiga to'liqroq to'lishi uchun kattalashtirildi
    font_miq  = _font(17, bold=True)
    font_tot  = _font(17, bold=True)

    dummy = Image.new("RGB", (10, 10))
    ddraw = ImageDraw.Draw(dummy)

    # DIQQAT: bu yerda tovar nomi to'g'ridan-to'g'ri (o'zgarishsiz) olinadi
    # — ilgari xitoy_nomi() chaqirilib, INVENTAR formatidagi nomni QAYTA
    # Xitoyning kamaytirilgan (masalan 0,9→0,85) formatiga aylantirib
    # yuborar edi. Bu "Kelgan yuklar" guruhiga yuboriladigan E'LON rasmi —
    # yolda_excel.py'da topilgan xuddi shu turdagi xato bu yerda ham bor
    # edi (qat'iy qoida: admin/filyal/guruh ko'radigan har qanday joy —
    # DOIM inventar formatida, xitoy_nomi() FAQAT Xitoyga yuboriladigan
    # "Yuklatish rejasi"da ishlatiladi).
    rows_prepared = []
    jami_miq = 0
    for _, r in grp.iterrows():
        tovar = str(r.get("Товар", ""))
        miq   = r.get("Миқдор", 0)
        jami_miq += miq if pd.notna(miq) else 0
        lines = _wrap(ddraw, tovar, font_name, COL_NAME_W - 24)
        h = max(ROW_H0, len(lines) * LINE_H + 16)
        rows_prepared.append((lines, miq, h))

    total_h = HDR_H + SUB_H + sum(r[2] for r in rows_prepared) + TOT_H + PAD * 2 + 6
    img  = Image.new("RGB", (IMG_W, int(total_h)), "#FFFFFF")
    draw = ImageDraw.Draw(img)

    y = PAD

    # ── Sarlavha — 🚛 konteyner nomidan OLDIN (kichikroq, matn bilan
    # proportsional o'lchamda), "Келди" (haqiqiy, taxminiy EMAS) sana bilan.
    # ✅ olib tashlandi — foydalanuvchi so'rovi bo'yicha (juda katta
    # ko'rinib, ortiqcha edi; "Келди" so'zining o'zi yetarli signal beradi).
    draw.rectangle([PAD, y, IMG_W - PAD, y + HDR_H], fill=f"#{CLR_ARRIVED_BG}")
    icon_h = 26   # matn balandligiga (~20px) yaqinroq, proporsional kichik
    cx = PAD + 12

    truck_icon = _emoji_icon("🚛", icon_h)
    if truck_icon:
        img.paste(truck_icon, (int(cx), int(y + (HDR_H - icon_h) / 2)), truck_icon)
        cx += truck_icon.width + 10

    draw.text((cx, y + HDR_H / 2), iso, font=font_hdr,
               fill=f"#{CLR_ARRIVED_TEXT}", anchor="lm")
    cx += draw.textlength(iso, font=font_hdr) + 14

    rest_text = f"│    Юкланган: {yukl_sana}    │    Келди: {keldi_sana}"
    if turi and turi != "STANDART":
        rest_text += f"    │    {turi}"
    draw.text((cx, y + HDR_H / 2), rest_text, font=font_hdr,
               fill=f"#{CLR_ARRIVED_TEXT}", anchor="lm")
    y += HDR_H

    # ── Ustun sarlavhalari ──────────────────────────────────────────────────
    draw.rectangle([PAD, y, IMG_W - PAD, y + SUB_H], fill=f"#{CLR_YOLDA_SUB}")
    x0 = PAD
    for label, w, align in [
        ("Товар",   COL_NAME_W, "l"),
        ("Миқдор",  COL_MIQ_W,  "c"),
    ]:
        if align == "l":
            draw.text((x0 + 12, y + SUB_H / 2), label, font=font_sub,
                       fill=f"#{CLR_COL_HDR_TEXT}", anchor="lm")
        else:
            draw.text((x0 + w / 2, y + SUB_H / 2), label, font=font_sub,
                       fill=f"#{CLR_COL_HDR_TEXT}", anchor="mm")
        x0 += w
    y += SUB_H

    # ── Tovar qatorlari ─────────────────────────────────────────────────────
    for i, (lines, miq, h) in enumerate(rows_prepared):
        bg = ROW_CLR_DARK if i % 2 == 0 else ROW_CLR_LIGHT
        draw.rectangle([PAD, y, IMG_W - PAD, y + h], fill=f"#{bg}")
        ty = y + h / 2 - (len(lines) - 1) * LINE_H / 2
        for ln in lines:
            draw.text((PAD + 12, ty), ln, font=font_name, fill="#222222", anchor="lm")
            ty += LINE_H
        draw.text((PAD + COL_NAME_W + COL_MIQ_W / 2, y + h / 2), str(miq),
                   font=font_miq, fill="#000000", anchor="mm")
        y += h

    # ── Jami qator — SONI EMAS, TONNA ko'rsatiladi (masalan "⚖️ 28 т") ──────
    draw.rectangle([PAD, y, IMG_W - PAD, y + TOT_H], fill=f"#{CLR_TOTAL_BG}")
    draw.text((PAD + COL_NAME_W - 12, y + TOT_H / 2), "ЖАМИ:",
               font=font_tot, fill="#000000", anchor="rm")

    tonna_text  = f"{jami_tonna} т"
    tot_cx      = PAD + COL_NAME_W + COL_MIQ_W / 2
    scale_icon  = _emoji_icon("⚖️", TOT_H - 18)
    if scale_icon:
        txt_w   = draw.textlength(tonna_text, font=font_tot)
        total_w = scale_icon.width + 8 + txt_w
        start_x = tot_cx - total_w / 2
        img.paste(scale_icon, (int(start_x), int(y + (TOT_H - scale_icon.height) / 2)), scale_icon)
        draw.text((start_x + scale_icon.width + 8, y + TOT_H / 2), tonna_text,
                   font=font_tot, fill="#000000", anchor="lm")
    else:
        # Emoji shrifti topilmasa — DejaVuSans'da ISHLAYDIGAN oddiy tarozi
        # belgisi ("⚖", VS16'siz) bilan almashtiriladi.
        draw.text((tot_cx, y + TOT_H / 2), f"⚖ {tonna_text}",
                   font=font_tot, fill="#000000", anchor="mm")
    y += TOT_H

    # ── Umumiy chegara ──────────────────────────────────────────────────────
    draw.rectangle([PAD, PAD, IMG_W - PAD, y], outline="#1A3A5C", width=2)

    bio = io.BytesIO()
    img.save(bio, format="PNG")
    bio.seek(0)
    return bio


# ── Standalone test ───────────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys
    iso_test = sys.argv[1] if len(sys.argv) > 1 else "TEST0000000"
    bio = generate_kelgan_rasm(iso_test)
    if bio:
        out = Path(__file__).parent / "chiqish" / f"kelgan_{iso_test}.png"
        out.write_bytes(bio.read())
        print(f"✅ Saqlandi: {out}")
    else:
        print("❌ Ma'lumot topilmadi")
