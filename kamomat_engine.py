"""
kamomat_engine.py — Kamomat tahlili motori
==========================================
Mantiq:
  - KRITIK + PAST tovarlarni aniqlash
  - Kunlik sotuv = min_zaxira / KELISH_KUNI  (common.py dan olinadi)
  - Zanjir simulyatsiyasi: konteynerlar sana bilan hisobga olinadi
  - Tartibli Excel: kategoriya -> o'lcham -> qalinlik -> uzunlik -> marka
  - Rangli Excel: har kategoriya o'z rang oilasida, juft/toq qatorlar

Bot.py import qiladi:
  from kamomat_engine import kamomat_stats_v2, kamomat_excel_v2
"""

import re, math, logging
import pandas as pd
from io import BytesIO
from pathlib import Path
from datetime import datetime

logger = logging.getLogger(__name__)

# ============================================================
# KONSTANTALAR — kelish kunini common.py dan olamiz
# ============================================================
from common import KELISH_KUNI
KUNLAR = KELISH_KUNI  # alias (eski kod o'zgarmasin)

CAT_ORDER = {
    "ТРУБА": 1, "ПРОФИЛЬ": 2, "ЛИСТ": 3, "ЛИСТ РУЛОН": 4,
    "БАЛАСИНА": 5, "СТОЙКА": 6, "СОККА": 7,
    "ШАР": 8, "ОТВОД": 9, "ЧАШКА": 10,
    "СОВУН": 11, "КУЗИКОРИН": 12, "БОШҚА": 13,
}

# Excel ranglari: h=header, a=toq qator, b=juft qator
CAT_COLORS = {
    "ТРУБА":      {"h": "9DC3E6", "a": "DEEBF7", "b": "BDD7EE"},
    "ПРОФИЛЬ":    {"h": "A9D18E", "a": "E2F0D9", "b": "C6E0B4"},
    "ЛИСТ":       {"h": "FFD966", "a": "FFF2CC", "b": "FFE699"},
    "ЛИСТ РУЛОН": {"h": "FFD966", "a": "FFF2CC", "b": "FFE699"},
    "БАЛАСИНА":   {"h": "C9B1D9", "a": "F0E5F7", "b": "E2CCEF"},
    "СТОЙКА":     {"h": "81C9C9", "a": "E0F2F2", "b": "C2E6E6"},
    "СОККА":      {"h": "F4B183", "a": "FCE4D6", "b": "F8CBAD"},
}
CAT_COLORS_DEF = {"h": "BFBFBF", "a": "F2F2F2", "b": "E8E8E8"}


# ============================================================
# SARALASH KALITI
# ============================================================
def tovar_sort_key(nom: str, kat: str) -> tuple:
    """
    Tovar nomidan saralash kalitini chiqaradi.
    Trubalarda: diametr -> qalinlik -> uzunlik -> marka
    Profillarda: en×boy -> qalinlik -> uzunlik -> marka
    Listlarda:  marka -> format -> qalinlik
    """
    cat_n = CAT_ORDER.get(kat, 99)
    n = str(nom)

    def _f(pat, default=99.0):
        m = re.search(pat, n)
        return float(m.group(1).replace(',', '.')) if m else default

    def _i(pat, default=999):
        m = re.search(pat, n)
        return int(m.group(1)) if m else default

    def _marka_n(text):
        m = re.search(r'\((\d+)\s*марка\)', text)
        return {"201": 1, "304": 2, "430": 3, "316": 4}.get(
            m.group(1) if m else "", 9)

    if kat == "ТРУБА":
        return (cat_n,
                _i(r'Ф-(\d+)'),
                _f(r'ст\s+(\d+[,.]?\d*)'),
                _f(r'\((\d+[,.]?\d*)\s*м\)'),
                _marka_n(n))

    elif kat == "ПРОФИЛЬ":
        s = re.search(r'(\d+)х(\d+)', n)
        w = int(s.group(1)) if s else 999
        h = int(s.group(2)) if s else 999
        return (cat_n, w, h,
                _f(r'ст\s+(\d+[,.]?\d*)'),
                _f(r'\((\d+[,.]?\d*)\s*м\)'),
                _marka_n(n))

    elif kat in ("ЛИСТ", "ЛИСТ РУЛОН"):
        fmt_m = re.search(r'\((\d+)х', n)
        return (cat_n,
                _marka_n(n),
                int(fmt_m.group(1)) if fmt_m else 9999,
                _f(r'Лист-\s*(\d+[,.]?\d*)'))

    return (cat_n, n)


# ============================================================
# ZANJIR SIMULYATSIYASI
# ============================================================
def zanjir_sim(qoldiq: float, min_z: float,
               konteynerlar: list) -> dict:
    """
    Qoldiq kuniga (min_z / 45) kamayadi.
    Har konteyner kelganda miqdor qo'shiladi.
    Qoldiq min_z dan past tushgan kun = uzilish_kun.

    konteynerlar: [(kun_qoldi: int, miqdor: float), ...]

    Qaytaradi:
      uzilish_kun  — None yoki bugundan necha kun (int)
      min_nuqta    — simulatsiya davomida eng past qoldiq
      taklif_A     — min ga qaytarish uchun buyurtma (50 ga yaxlit)
      taklif_B     — zanjirni to'liq ta'minlash uchun (50 ga yaxlit)
      taklif       — max(A, B)
      xavf         — 'KRITIK' | 'PAST' | 'NORMA' | 'MEYOR_YOQ'
    """
    EMPTY = dict(uzilish_kun=None, min_nuqta=qoldiq,
                 taklif_A=0, taklif_B=0, taklif=0, xavf="MEYOR_YOQ")
    if min_z <= 0:
        return EMPTY

    kunlik = min_z / float(KUNLAR)
    kont   = sorted(konteynerlar, key=lambda x: x[0])

    joriy       = float(qoldiq)
    joriy_kun   = 0.0
    uzilish_kun = None
    min_nuqta   = joriy

    for kun_q, miqdor in kont:
        if kun_q <= joriy_kun:
            joriy += miqdor
            continue
        gap         = kun_q - joriy_kun
        joriy_oldin = joriy
        joriy      -= kunlik * gap
        min_nuqta   = min(min_nuqta, joriy)

        if joriy < min_z and uzilish_kun is None:
            dtm = max(0.0, (joriy_oldin - min_z) / kunlik)
            uzilish_kun = int(joriy_kun + dtm)

        joriy    += miqdor
        joriy_kun = kun_q

    # Oxirgi konteynerdan keyin KELISH_KUNI kun tekshiruv
    joriy_oxir = joriy - kunlik * KUNLAR
    min_nuqta  = min(min_nuqta, joriy_oxir)
    if joriy_oxir < min_z and uzilish_kun is None:
        dtm = max(0.0, (joriy - min_z) / kunlik)
        uzilish_kun = int(joriy_kun + dtm)

    # Taklif hisoblash
    yolda_jami = sum(m for _, m in kont)

    def _50(x: float) -> int:
        return int(math.ceil(x / 50)) * 50 if x > 0 else 0

    taklif_A = max(0.0, min_z - (qoldiq + yolda_jami))
    taklif_B = max(0.0, min_z - min_nuqta) if min_nuqta < min_z else 0.0

    # Xavf darajasi
    if uzilish_kun is not None and uzilish_kun <= KELISH_KUNI:
        xavf = "KRITIK"
    elif uzilish_kun is not None:
        xavf = "PAST"
    elif qoldiq < min_z * 1.5:
        xavf = "PAST"
    else:
        xavf = "NORMA"

    return dict(
        uzilish_kun=uzilish_kun,
        min_nuqta=int(round(min_nuqta)),
        taklif_A=_50(taklif_A),
        taklif_B=_50(taklif_B),
        taklif=_50(max(taklif_A, taklif_B)),
        xavf=xavf,
    )


# ============================================================
# KAMOMAT STATISTIKASI
# ============================================================
def kamomat_stats_v2(data_file: Path, kanal: str,
                     buyurtma_yuklash_fn) -> dict:
    """
    KRITIK + PAST tovarlar sonini va buyurtma holatini qaytaradi.
    Qaytaradi: {n, kritik, past, b (berildi), p (pending)}
    """
    try:
        df = pd.read_excel(data_file, sheet_name="Инвентар")
        if "Тур" in df.columns:
            df = (df[df["Тур"] == "ЦЕХ🏭"] if kanal == "sex"
                  else df[df["Тур"] != "ЦЕХ🏭"])

        kritik = int((df["Холат"] == "🔴 КРИТИК").sum())
        past   = int((df["Холат"] == "🟡 ПАСТ").sum())
        jami   = kritik + past

        if jami == 0:
            return {"n": 0, "kritik": 0, "past": 0, "b": 0, "p": 0}

        buy     = buyurtma_yuklash_fn(kanal)
        ordered = {i["tovar"] for i in buy.get("buyurtmalar", [])} if buy else set()
        kamomat = df[df["Холат"].isin(["🔴 КРИТИК", "🟡 ПАСТ"])]
        b       = sum(1 for t in kamomat.get("Товар", pd.Series()).tolist()
                      if t in ordered)
        return {"n": jami, "kritik": kritik, "past": past,
                "b": b, "p": jami - b}
    except Exception as e:
        logger.error(f"kamomat_stats_v2: {e}")
        return {"n": 0, "kritik": 0, "past": 0, "b": 0, "p": 0}


# ============================================================
# EXCEL GENERATSIYASI
# ============================================================
def kamomat_excel_v2(data_file: Path, kanal: str,
                     lang: str, buyurtma_yuklash_fn) -> BytesIO | None:
    """
    Tartibli, rangli kamomat Excel.
    - KRITIK + PAST tovarlar
    - Kategoriya bo'yicha saralanadi (TRUБА -> PROFIL -> LIST ...)
    - Har kategoriya o'z rang oilasida (juft/toq qatorlar)
    - Kategoriya separator qatori (ajralib turadi)
    - Ustunlar: №, Товар, Холат, Қолдиқ, Йўлда, Мин_Захира,
                Кун_Хавф, Буюртма_Ҳолати, Таклиф_Миқдор
    """
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

    # ── Ma'lumot o'qish ──────────────────────────────────────
    try:
        inv = pd.read_excel(data_file, sheet_name="Инвентар")
        for col in ["Қолдиқ", "Мин_Захира", "Йўлда_Жами", "Фарқ"]:
            if col in inv.columns:
                inv[col] = pd.to_numeric(inv[col], errors="coerce").fillna(0)
    except Exception as e:
        logger.error(f"kamomat_excel_v2 inv: {e}")
        return None

    # Kanal filtri
    if "Тур" in inv.columns:
        inv = (inv[inv["Тур"] == "ЦЕХ🏭"] if kanal == "sex"
               else inv[inv["Тур"] != "ЦЕХ🏭"])

    df = inv[inv["Холат"].isin(["🔴 КРИТИК", "🟡 ПАСТ"])].copy()
    if df.empty:
        return None

    # ── Konteyner ma'lumotlari ────────────────────────────────
    kont_map: dict[str, list] = {}
    try:
        kont = pd.read_excel(data_file, sheet_name="Контейнерлар")
        kont = kont[kont["Холат"] != "КЕЛДИ ✅"].copy()
        for col in ["Кун_Қолди", "Миқдор"]:
            if col in kont.columns:
                kont[col] = pd.to_numeric(kont[col], errors="coerce").fillna(0)
        for _, r in kont.iterrows():
            tovar = str(r.get("Товар", ""))
            kq    = float(r.get("Кун_Қолди", 0))
            mq    = float(r.get("Миқдор", 0))
            if tovar and mq > 0:
                kont_map.setdefault(tovar, []).append((kq, mq))
    except Exception as e:
        logger.warning(f"Konteyner o'qilmadi: {e}")

    # ── Buyurtma holati ────────────────────────────────────────
    buy     = buyurtma_yuklash_fn(kanal)
    ordered = {i["tovar"] for i in buy.get("buyurtmalar", [])} if buy else set()
    berildi = "Берилди ✅" if lang == "cyr" else "Berildi ✅"
    kutilmq = "Кутилмоқда ⏳" if lang == "cyr" else "Kutilmoqda ⏳"

    # ── Saralash ───────────────────────────────────────────────
    if "Категория" not in df.columns:
        df["Категория"] = "БОШҚА"
    df["_c"]   = df["Категория"].apply(lambda x: CAT_ORDER.get(str(x), 99))
    df["_s"]   = df.apply(lambda r: tovar_sort_key(
                    str(r.get("Товар", "")), str(r.get("Категория", ""))), axis=1)
    df = df.sort_values(["_c", "_s"]).reset_index(drop=True)

    # ── Zanjir simulyatsiyasi ──────────────────────────────────
    sims = []
    for _, row in df.iterrows():
        tovar  = str(row.get("Товар", ""))
        qoldiq = float(row.get("Қолдиқ", 0))
        min_z  = float(row.get("Мин_Захира", 0))
        kont_l = kont_map.get(tovar, [])
        sims.append(zanjir_sim(qoldiq, min_z, kont_l))

    # ── Excel ──────────────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Камомат" if lang == "cyr" else "Kamomat"

    # Sarlavhalar
    if lang == "cyr":
        hdrs = ["№", "Товар", "Холат", "Қолдиқ", "Йўлда_Жами",
                "Мин_Захира", "Кун_Хавф", "Буюртма_Ҳолати", "Таклиф_Миқдор"]
    else:
        hdrs = ["№", "Tovar", "Holat", "Qoldiq", "Yolda_Jami",
                "Min_Zaxira", "Kun_Xavf", "Buyurtma_Holati", "Taklif_Miqdor"]

    col_w = [5, 46, 13, 11, 12, 13, 11, 19, 16]
    NCOL  = len(hdrs)

    # Stil yordamchilari
    def fill(hex_: str):
        return PatternFill("solid", fgColor=hex_)

    def font(sz=10, bold=False, color="000000"):
        return Font(size=sz, bold=bold, color=color)

    def aln(h="left", v="center", wrap=False):
        return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

    thin = Side(style="thin", color="D0D0D0")
    border_thin = Border(bottom=thin)

    # Header qatori
    ws.append(hdrs)
    ws.row_dimensions[1].height = 28
    hdr_fill = fill("1F4E79")
    hdr_font = font(11, True, "FFFFFF")
    hdr_aln  = aln("center", "center", True)
    for i, cell in enumerate(ws[1], 1):
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.alignment = hdr_aln
        ws.column_dimensions[cell.column_letter].width = col_w[i - 1]

    ws.freeze_panes = "A2"

    # ── Qatorlar ──────────────────────────────────────────────
    cur_kat    = None
    kat_cnt    = 0    # har kategoriya ichida qator hisobi
    excel_row  = 2
    n          = 0

    for idx, row in df.iterrows():
        kat    = str(row.get("Категория", "БОШҚА"))
        tovar  = str(row.get("Товар", ""))
        holat  = str(row.get("Холат", ""))
        qoldiq = int(row.get("Қолдиқ", 0))
        yolda  = int(row.get("Йўлда_Жами", 0))
        min_z  = int(row.get("Мин_Захира", 0))
        sim    = sims[idx]

        # Kategoriya separator
        if kat != cur_kat:
            cur_kat = kat
            kat_cnt = 0
            colors  = CAT_COLORS.get(kat, CAT_COLORS_DEF)

            ws.append([kat] + [""] * (NCOL - 1))
            ws.merge_cells(
                start_row=excel_row, start_column=1,
                end_row=excel_row, end_column=NCOL
            )
            sep_cell           = ws.cell(row=excel_row, column=1)
            sep_cell.value     = f"  {kat}"
            sep_cell.fill      = fill(colors["h"])
            sep_cell.font      = font(11, True, "1F1F1F")
            sep_cell.alignment = aln("left", "center")
            ws.row_dimensions[excel_row].height = 22
            excel_row += 1

        # Ma'lumot qatori
        n      += 1
        kat_cnt += 1
        colors  = CAT_COLORS.get(kat, CAT_COLORS_DEF)
        row_clr = colors["a"] if kat_cnt % 2 == 1 else colors["b"]

        # Kун_Хавф qiymati
        uzilish = sim.get("uzilish_kun")
        if uzilish is None:
            kun_xavf = "—"
        elif uzilish <= 0:
            kun_xavf = "XOZIR ❗" if lang == "lat" else "ҲОЗИР ❗"
        else:
            kun_xavf = f"{uzilish} kun" if lang == "lat" else f"{uzilish} кун"

        taklif       = sim.get("taklif", 0)
        b_holat      = berildi if tovar in ordered else kutilmq

        ws.append([
            n, tovar, holat, qoldiq, yolda,
            min_z, kun_xavf, b_holat,
            taklif if taklif > 0 else ""
        ])

        row_fill = fill(row_clr)
        for ci, cell in enumerate(ws[excel_row], 1):
            cell.fill   = row_fill
            cell.font   = font(10)
            cell.border = border_thin

            if ci == 1:                       # №
                cell.alignment = aln("center")
            elif ci == 2:                     # Товар
                cell.alignment = aln("left")
            elif ci == 3:                     # Холат
                cell.alignment = aln("center")
                if "КРИТИК" in holat:
                    cell.font = font(10, True, "C00000")
                elif "ПАСТ" in holat:
                    cell.font = font(10, True, "7B3F00")
            elif ci in (4, 5, 6, 9):         # Raqamlar
                cell.alignment = aln("right")
            elif ci == 7:                     # Кун_Хавф
                cell.alignment = aln("center")
                if uzilish is not None and uzilish <= 15:
                    cell.font = font(10, True, "C00000")
                elif uzilish is not None and uzilish <= 30:
                    cell.font = font(10, False, "7B3F00")
            else:
                cell.alignment = aln("center")

        ws.row_dimensions[excel_row].height = 17
        excel_row += 1

    # Umumiy info (oxirda)
    excel_row += 1
    info_text = (
        f"Sana: {datetime.now().strftime('%d.%m.%Y %H:%M')}  |  "
        f"Kanal: {kanal}  |  "
        f"Jami: {n} ta  |  "
        f"Kunlik = min/{KELISH_KUNI}"
    )
    ws.append([info_text])
    ws.merge_cells(
        start_row=excel_row, start_column=1,
        end_row=excel_row, end_column=NCOL
    )
    info_cell           = ws.cell(row=excel_row, column=1)
    info_cell.font      = font(9, False, "808080")
    info_cell.alignment = aln("left")

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


# ============================================================
# ZANJIR GRAFIK
# ============================================================

# ============================================================
# GRAFIK
# ============================================================

# ============================================================
# GRAFIK
# ============================================================
def grafik_chiz(tovar: str, qoldiq: float, min_z: float,
                konteynerlar: list, kunlar: int = 0) -> "BytesIO | None":
    """
    Piecewise-linear stock trace:
      - qoldiq kunlik ravishda pasayadi
      - konteyner kelgan kunda qoldiq keskin ko'tariladi
      - har bir konteyner uchun alohida label (rotated, linea bo'yida)
      - bir kunda bir nechta: bitta chiziq, labellar stacked
    """
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        import matplotlib.ticker as ticker
        from collections import defaultdict
    except ImportError:
        return None

    kunlik = min_z / float(KELISH_KUNI) if min_z > 0 else 1.0

    # ── Konteynerlarni kun bo'yicha guruhlash ──────────────────
    kont_sorted = sorted(konteynerlar, key=lambda x: x[0])
    kont_by_day = defaultdict(list)
    for kq, mq in kont_sorted:
        d = max(0, int(round(kq)))
        kont_by_day[d].append(float(mq))

    # ── X oralig'i ─────────────────────────────────────────────
    max_day = max(kont_by_day.keys()) if kont_by_day else 0
    if kunlar == 0:
        kunlar = max(max_day + KELISH_KUNI + 10, 90)
    kunlar = min(kunlar, 200)

    # ── Piecewise-linear trace ─────────────────────────────────
    # Har bir segment: linear tushish, keyin konteynerda sakrash
    tx, ty = [0], [float(qoldiq)]
    cur = float(qoldiq)
    prev_d = 0

    for d in sorted(kont_by_day.keys()):
        if d > kunlar:
            continue
        # tushish davri
        elapsed = d - prev_d
        cur_before = max(cur - kunlik * elapsed, 0)
        if d > prev_d:
            tx.append(d)
            ty.append(cur_before)
        # sakrash
        cur = cur_before + sum(kont_by_day[d])
        tx.append(d)
        ty.append(cur)
        prev_d = d

    # oxirgi nuqtagacha tushish
    remaining = kunlar - prev_d
    end_y = max(cur - kunlik * remaining, 0)
    tx.append(kunlar)
    ty.append(end_y)

    # ── Y o'qi chegarasi ───────────────────────────────────────
    y_max = max(max(ty) * 1.10, min_z * 1.35)

    # ── Uzilish kuni (chiziq min_z dan pastga tushgan kun) ─────
    uzilish_day = None
    for i in range(1, len(tx)):
        x0, y0 = tx[i-1], ty[i-1]
        x1, y1 = tx[i], ty[i]
        if y0 >= min_z and y1 < min_z and x1 > x0:
            # interpolatsiya
            frac = (y0 - min_z) / (y0 - y1)
            uzilish_day = int(x0 + frac * (x1 - x0))
            break
        elif y0 < min_z and uzilish_day is None:
            uzilish_day = int(x0)
            break

    # ── Ranglar ────────────────────────────────────────────────
    BG    = "#0F1923"
    GRID  = "#1A2B3C"
    C_STK = "#3498DB"
    C_MIN = "#E74C3C"
    C_KNT = "#2ECC71"
    C_UZL = "#E74C3C"

    fig, ax = plt.subplots(figsize=(13, 5))
    fig.patch.set_facecolor(BG)
    ax.set_facecolor(BG)

    # ── Uzilish zonasi ─────────────────────────────────────────
    if uzilish_day is not None and uzilish_day < kunlar:
        ax.axvspan(uzilish_day, kunlar, alpha=0.12, color=C_UZL, zorder=1)

    # ── Min zaxira ─────────────────────────────────────────────
    ax.axhline(y=min_z, color=C_MIN, linewidth=2.0,
               linestyle="--", alpha=1.0, zorder=3)
    ax.text(kunlar * 0.01, min_z * 1.02,
            "Min: {:,}".format(int(min_z)).replace(",", " "),
            color=C_MIN, fontsize=8, fontweight="bold", va="bottom", zorder=6)

    # ── Qoldiq chizig'i ────────────────────────────────────────
    ax.plot(tx, ty, color=C_STK, linewidth=2.2, zorder=4)
    ax.plot(0, qoldiq, "o", color=C_STK, markersize=5, zorder=5)

    # ── Konteyner chiziqlari + rotated labellar ────────────────
    total_kont = sum(len(v) for v in kont_by_day.values())
    sorted_days = sorted(kont_by_day.keys())

    for d in sorted_days:
        if d > kunlar:
            continue
        containers = kont_by_day[d]

        # qoldiq darajasi d kunda (sakrashdan OLDIN)
        elapsed = d - (sorted_days[sorted_days.index(d) - 1]
                       if sorted_days.index(d) > 0 else 0)
        # tx/ty dan d nuqtasini topamiz (sakrashdan oldingi qiymat)
        y_before = 0
        for i, x in enumerate(tx):
            if x == d:
                y_before = ty[i]   # sakrashdan oldingi qiymat
                break
            if x > d:
                # interpolatsiya
                x0, y0 = tx[i-1], ty[i-1]
                x1, y1 = tx[i], ty[i]
                if x1 > x0:
                    y_before = y0 + (y1 - y0) * (d - x0) / (x1 - x0)
                break

        ax.axvline(x=d, color=C_KNT, linewidth=1.8, alpha=0.9, zorder=4)

        # Har bir konteyner: label chiziq ustida, boshqa y darajasida
        # rotation=90 → matn vertikal, har biri o'z qoldiq segmentida
        cum_y = y_before
        for j, mq in enumerate(containers):
            label_y = cum_y + mq * 0.5   # segmentning o'rtasi
            qty_str = "+{:,}".format(int(mq)).replace(",", " ")
            ax.text(d, label_y, qty_str,
                    color=C_KNT, fontsize=8, fontweight="bold",
                    rotation=90, ha="center", va="center",
                    zorder=6,
                    bbox=dict(boxstyle="round,pad=0.1",
                              fc="#0F1923", ec="none", alpha=0.7))
            cum_y += mq

    # ── Uzilish belgisi ────────────────────────────────────────
    if uzilish_day is not None and 0 < uzilish_day < kunlar:
        ax.axvline(x=uzilish_day, color=C_UZL, linewidth=1.5,
                   linestyle=":", alpha=0.9, zorder=5)
        ax.text(uzilish_day + kunlar * 0.005, min_z * 0.5,
                "UZILISH {}k".format(uzilish_day),
                color=C_UZL, fontsize=8, fontweight="bold",
                va="center", zorder=6)

    # ── O'qlar — real sanalar bilan ────────────────────────────
    import calendar as _cal
    from datetime import date as _date, timedelta as _td

    _bugun = _date.today()

    # Har oyning 10, 20 va oxirgi kuni da tick
    _ticks, _labels = [], []
    _prev_month = None
    for _off in range(kunlar + 1):
        _d = _bugun + _td(days=_off)
        _last = _cal.monthrange(_d.year, _d.month)[1]
        if _d.day in (10, 20, _last):
            _ticks.append(_off)
            # Oy o'zgarganda nomi ham ko'rsatiladi
            _ay = ["Yan","Fev","Mar","Apr","May","Iyun",
                   "Iyul","Avg","Sen","Okt","Noy","Dek"][_d.month - 1]
            if _d.month != _prev_month:
                _labels.append(f"{_d.day}\n{_ay}")
                _prev_month = _d.month
            else:
                _labels.append(str(_d.day))

    ax.set_xlim(0, kunlar)
    ax.set_xticks(_ticks)
    ax.set_xticklabels(_labels, fontsize=7, color="#8899AA")
    ax.tick_params(axis="x", which="major", length=4, color="#334455")

    ax.set_ylim(0, y_max)
    ax.yaxis.set_major_formatter(ticker.FuncFormatter(
        lambda x, _: "{:,}".format(int(x)).replace(",", " ")))

    ax2 = ax.twinx()
    ax2.set_ylim(0, y_max)
    ax2.set_facecolor(BG)
    ax2.tick_params(colors="#8899AA", labelsize=8)
    ax2.yaxis.set_major_formatter(ticker.FuncFormatter(
        lambda x, _: "{:.0f}".format(x / kunlik) if kunlik > 0 else "0"))
    ax2.set_ylabel("Kunlar", color="#8899AA", fontsize=8)
    for sp in ax2.spines.values():
        sp.set_color(GRID)

    ax.set_xlabel(
        "Kunlik: {:,} dona".format(int(kunlik)).replace(",", " "),
        color="#8899AA", fontsize=9)
    ax.set_ylabel("Qoldiq (dona)", color="#8899AA", fontsize=9)
    ax.tick_params(colors="#8899AA", labelsize=8)
    for sp in ax.spines.values():
        sp.set_color(GRID)
    ax.grid(True, alpha=0.15, color=GRID, linestyle="-")

    short = tovar if len(tovar) <= 58 else tovar[:55] + "..."
    ax.set_title(short, color="#FFFFFF", fontsize=10, pad=10)

    handles = [
        plt.Line2D([0], [0], color=C_STK, linewidth=2.2, label="Qoldiq"),
        plt.Line2D([0], [0], color=C_MIN, linewidth=2.0,
                   linestyle="--", label="Min zaxira"),
        plt.Line2D([0], [0], color=C_KNT, linewidth=1.8,
                   label="Konteyner ({} ta)".format(total_kont)),
    ]
    ax.legend(handles=handles, facecolor="#1A2A3A",
              labelcolor="#CCDDEE", fontsize=8.5,
              loc="upper right", framealpha=0.9)

    plt.tight_layout(pad=1.0)

    bio = BytesIO()
    plt.savefig(bio, format="png", dpi=120,
                bbox_inches="tight", facecolor=BG)
    plt.close(fig)
    bio.seek(0)
    return bio
