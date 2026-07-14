"""
NEJAVIYKA INVENTAR TAHLIL v4 — Power BI uchun
===============================================
v4 da o'zgarishlar:
  - min_cache tizimi olib tashlandi
  - Min_Zaxira.xlsx dan o'qiydi (siz qo'lda tahrirlaysiz)
  - min_cache.xlsx kerak emas

Arxitektura:
  common.py             ← barcha umumiy funksiyalar
  yarat_min_zaxira.py   ← Min_Zaxira.xlsx ni bir marta yaratadi
  NEJAVIYKA_v3.py       ← asosiy tahlil (shu fayl)
===============================================
"""

import sys
import re
import pandas as pd
import numpy as np
import glob
import os
from datetime import datetime, timedelta
from pathlib import Path
import warnings

warnings.filterwarnings('ignore')

# --- Umumiy funksiyalar ---
from common import (
    normalize_product_name,
    get_category, get_marka, get_category_with_marka, get_category_order,
    get_himoya_foiz, yaxlitla_50, hisobla_min_zaxira, min_dan_kunlik_chiqar,
    load_qoldiq_file, keraksizmi, fayl_sanasi,
    YOLDA_KUN, FAST_KUN, M12_KUN,
)


# ============================================================
# YO'LLAR
# ============================================================

SCRIPT_DIR        = Path(__file__).resolve().parent
TARIX_FOLDER      = SCRIPT_DIR / 'tarix'
MIN_ZAXIRA_FILE   = SCRIPT_DIR / 'Minimal_zaxiralar' / 'Min_Zaxira.xlsx'
KONTEYNER_FOLDER       = SCRIPT_DIR / 'konteynerlar'
KONTEYNER_XITOY_FOLDER = SCRIPT_DIR / 'konteynerlar' / 'xitoy_parsed'
OUTPUT_FILE       = SCRIPT_DIR / 'chiqish' / 'NEJAVIYKA_POWER_BI.xlsx'

# ============================================================
# BOSHLASH
# ============================================================

print("=" * 100)
print("NEJAVIYKA INVENTAR TAHLIL v3 — Dinamik Min_Zaxira + Ishlab Chiqarish")
print(f"Ishga tushdi: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
print("=" * 100)


# ============================================================
# 1. QOLDIQNI YUKLASH
# ============================================================

print("\n1. QOLDIQNI YUKLASH...")

tarix_fayllar = sorted(glob.glob(str(TARIX_FOLDER / '*.xlsx')), key=fayl_sanasi)
if not tarix_fayllar:
    print("❌ Tarix papkasida fayl yo'q! Dastur to'xtatildi.")
    sys.exit(1)

QOLDIQ_FILE = Path(tarix_fayllar[-1])
print(f"  📅 Bugungi qoldiq fayli: {QOLDIQ_FILE.name}")

qoldiq = load_qoldiq_file(str(QOLDIQ_FILE))
print(f"  ✅ {len(qoldiq):,} mahsulot yuklandi")


# ============================================================
# 2. MIN_ZAXIRA YUKLASH (Min_Zaxira.xlsx)
# ============================================================

print("\n2. MIN_ZAXIRA YUKLANMOQDA...")

try:
    mz_raw = pd.read_excel(MIN_ZAXIRA_FILE, sheet_name='Min_Zaxira')
    # Ustunlar: №, Товар, Категория, Йиллик_Сотув, Кунлик_Истеъмол, Мин_Захира, (Цех_Захира)

    # Ustun nomlarini ANIQ matnga emas, KALIT SO'ZGA qarab topamiz —
    # chunki Excel'da sarlavhalar qo'lda qayta yozilganda harflar
    # (kirill/lotin) biroz boshqacha bo'lib qolishi mumkin.
    def _find_col(cols, include, exclude=()):
        for c in cols:
            cl = str(c).lower()
            if any(k in cl for k in include) and not any(k in cl for k in exclude):
                return c
        return None

    cols = list(mz_raw.columns)
    tovar_col  = _find_col(cols, ['товар', 'tovar'])
    kat_col    = _find_col(cols, ['категор', 'kategor'])
    kunlik_col = _find_col(cols, ['кунлик', 'kunlik'])
    cex_col    = _find_col(cols, ['цех', 'tsex', 'cex'])
    sotuv_col  = _find_col(cols, ['захира', 'zaxira', 'мин', 'min'],
                            exclude=['цех', 'tsex', 'cex'])

    if not tovar_col or not sotuv_col:
        raise KeyError(
            f"Kerakli ustunlar topilmadi. Mavjud ustunlar: {cols} | "
            f"tovar={tovar_col}, kategoriya={kat_col}, kunlik={kunlik_col}, "
            f"sotuv_min={sotuv_col}, cex_min={cex_col}"
        )

    mz_raw = mz_raw[mz_raw[tovar_col].notna()].copy()

    mz_raw['Sotuv_Min'] = pd.to_numeric(mz_raw[sotuv_col], errors='coerce').fillna(0)
    mz_raw['Kunlik_Istemol'] = (
        pd.to_numeric(mz_raw[kunlik_col], errors='coerce').fillna(0)
        if kunlik_col else 0
    )
    mz_raw['Kategoriya'] = mz_raw[kat_col] if kat_col else 'БОШҚА'

    if cex_col:
        mz_raw['Cex_Min'] = pd.to_numeric(mz_raw[cex_col], errors='coerce').fillna(0)
        print(f"  ✅ Цех ustuni topildi: '{cex_col}'")
    else:
        mz_raw['Cex_Min'] = 0
        print("  ⚠️  Цех_Захира ustuni topilmadi — barcha Cex_Min = 0")

    mz_raw['Mahsulot_Normalized'] = mz_raw[tovar_col].apply(normalize_product_name)
    min_zaxira = mz_raw[['Mahsulot_Normalized', 'Sotuv_Min', 'Kunlik_Istemol', 'Kategoriya', 'Cex_Min']].copy()
    print(f"  ✅ {len(min_zaxira):,} ta tovar min zaxirasi yuklandi")
    print(f"  ✅ {(min_zaxira['Cex_Min'] > 0).sum()} ta tovarda Цех_Захира > 0 (Tsex mahsulotlari)")
except FileNotFoundError:
    print(f"  ❌ Min_Zaxira.xlsx topilmadi: {MIN_ZAXIRA_FILE}")
    print(f"     Avval 'yarat_min_zaxira.py' ni ishga tushiring!")
    sys.exit(1)
except Exception as e:
    print(f"  ❌ Min_Zaxira.xlsx xatosi: {e}")
    sys.exit(1)


# ============================================================
# 4. KONTEYNERLARNI YUKLASH
# ============================================================

print("\n4. KONTEYNERLARNI YUKLASH...")

def _konteyner_12m_mi(mahsulot_qatorlari: list) -> bool:
    """
    "12 metrlik" konteyner (2026-07-08 qo'shildi) — ISO raqamisiz,
    ko'pincha mashina-raqami (masalan "ME5312") bilan yozilgan, tarkibida
    ASOSAN 6 metrlik Труба/Профиль bo'lgan yetkazma. Fayl nomidan EMAS,
    balki tarkibidagi tovar nomlaridan aniqlanadi — chunki nomi ba'zida
    mashina raqami, ba'zida haqiqiy konteyner raqami bo'lib qolishi mumkin
    (Huzayfa aniqligi bilan tasdiqladi). Qoida: tovar qatorlarining
    YARMIDAN KO'PI nomida aniq "(6 м)" bo'lsa → 12 metrlik.

    Xato yoki noaniq holatda XAVFSIZ False qaytaradi (ya'ni oddiy 55
    kunlik hisoblanadi) — noto'g'ri tezlashtirib yuborishdan ko'ra,
    noto'g'ri sekinlashtirib qo'yish xavfsizroq (admin "Кечикди"
    ogohlantirishini ko'rib, qo'lda tekshirib chiqishi mumkin).
    """
    try:
        if not mahsulot_qatorlari:
            return False
        olti_m_soni = sum(
            1 for m in mahsulot_qatorlari
            if re.search(r'\(6\s*м\)', str(m))
        )
        return olti_m_soni > len(mahsulot_qatorlari) / 2
    except Exception:
        return False


def _parse_konteyner_fayli(file_path: str) -> list[dict]:
    """
    Bir konteyner faylini o'qib, tovar qatorlari ro'yxatini qaytaradi.
    Fayl nomi formatlar: 'K001_15.04.2025.xlsx' yoki 'F_K001_15.04.2025.xlsx'
    _D suffiks: yetib kelgan konteyner.

    Yetib kelish muddati UCH XIL bo'lishi mumkin (2026-07-08):
      • FAST_KUN (fayl nomi "F_" bilan boshlansa) — eng ustuvor, tarkibidan
        qat'iy nazar.
      • M12_KUN ("12 metrlik" — tarkibida ko'pchilik tovar 6 metrlik,
        qarang _konteyner_12m_mi()).
      • YOLDA_KUN — aks holda, oddiy dengiz konteyneri.
    """
    filename = os.path.basename(file_path)
    name     = filename.replace('.xlsx', '')

    # FAST konteyner (tezkor)
    is_fast = name.startswith('F_')
    if is_fast:
        name = name[2:]

    # Yetib kelgan konteyner
    is_done = name.endswith('_D')
    if is_done:
        name = name[:-2]

    parts = name.split('_')
    if len(parts) < 2:
        print(f"  ⚠️  Fayl nomi formati noto'g'ri (o'tkazib yuborildi): {filename}")
        return []

    konteyner_raqam = parts[0]
    date_str        = parts[1]

    departure_date = None
    for fmt in ['%d.%m.%Y', '%d_%m_%Y', '%Y-%m-%d']:
        try:
            departure_date = datetime.strptime(date_str, fmt)
            break
        except ValueError:
            continue

    if departure_date is None:
        print(f"  ⚠️  Sana aniqlanmadi (o'tkazib yuborildi): {filename}")
        return []

    rows = []
    try:
        df           = pd.read_excel(file_path)
        mahsulot_col = df.columns[1]

        # Birinchi raqamli ustunni miqdor sifatida olish
        miqdor_col = next(
            (col for col in df.columns[2:] if pd.api.types.is_numeric_dtype(df[col])),
            df.columns[2]
        )

        # "Vazn_kg" ustuni (agar mavjud bo'lsa) — konteyner_qosh.py
        # tomonidan BIR MARTA (qo'shilayotgan paytda) hisoblab yozib
        # qo'yilgan og'irlik. Mavjud bo'lsa shu qiymat ishlatiladi — bu
        # yerda yoki keyinroq (yolda_excel.py'da) tovar nomidan qayta
        # hisoblashga hojat qolmaydi. Eski (bu ustun yo'q) fayllar uchun
        # None qoladi — yolda_excel.py bunday holatda tovar nomidan
        # taxminiy hisoblaydi (orqaga moslashuvchan fallback).
        vazn_col = next(
            (c for c in df.columns if str(c).strip().lower() in
             ('vazn_kg', 'вазн_кг', 'vazn (kg)', 'vazn')),
            None
        )

        # 12-metrlik aniqlash uchun avval barcha tovar nomlarini yig'ib
        # olamiz (transit_days'ni belgilashdan OLDIN kerak).
        mahsulot_qatorlari = [
            str(m).strip() for m in df[mahsulot_col].tolist()
            if pd.notna(m) and str(m).strip()
        ]
        is_12m = (not is_fast) and _konteyner_12m_mi(mahsulot_qatorlari)

        if is_fast:
            transit_days = FAST_KUN
            turi         = 'FAST'
        elif is_12m:
            transit_days = M12_KUN
            turi         = '12M'
        else:
            transit_days = YOLDA_KUN
            turi         = 'STANDART'

        arrival_date = departure_date + timedelta(days=transit_days)
        bugun        = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

        # Kechikish hisobi
        if not is_done and arrival_date < bugun:
            kechikkan_kun  = (bugun - arrival_date).days
            current_status = 'КЕЧИКДИ ⚠️'
        elif is_done:
            kechikkan_kun  = 0
            current_status = 'КЕЛДИ ✅'
        else:
            kechikkan_kun  = 0
            current_status = 'ЙЎЛДА 🚢'

        for _, row in df.iterrows():
            mahsulot = row[mahsulot_col]
            if pd.isna(mahsulot):
                continue
            mahsulot_str = str(mahsulot).strip()
            if not mahsulot_str:
                continue
            # Sarlavha qatorlarini o'tkazib yuborish
            if any(x in mahsulot_str.lower() for x in ['жами', 'jami', 'total', '№']):
                continue

            try:
                qty = float(row[miqdor_col])
                if qty > 0:
                    vazn_kg = None
                    if vazn_col is not None:
                        try:
                            v = row[vazn_col]
                            vazn_kg = float(v) if pd.notna(v) and v != '' else None
                        except (ValueError, TypeError):
                            vazn_kg = None
                    rows.append({
                        'Konteyner_Raqam':    konteyner_raqam,
                        'Mahsulot_Normalized': normalize_product_name(mahsulot_str),
                        'Konteyner_Miqdor':   qty,
                        'Konteyner_Vazn_Kg':  vazn_kg,
                        'Chiqish_Sana':       departure_date,
                        'Kelish_Sana':        arrival_date,
                        'Transit_Kun':        transit_days,
                        'Konteyner_Turi':     turi,
                        'Holat':              current_status,
                        'Kechikish_Kun':      kechikkan_kun,
                    })
            except (ValueError, TypeError):
                continue
    except Exception as e:
        print(f"  ❌ O'qib bo'lmadi ({filename}): {e}")

    return rows


all_containers = []
_kont_fayllar = (
    sorted(glob.glob(str(KONTEYNER_FOLDER / '*.xlsx'))) +
    sorted(glob.glob(str(KONTEYNER_XITOY_FOLDER / '*.xlsx')))
)
for file_path in _kont_fayllar:
    rows = _parse_konteyner_fayli(file_path)
    if rows:
        print(f"  ✅ {os.path.basename(file_path)} — {len(rows)} qator")
    all_containers.extend(rows)

if all_containers:
    containers = pd.DataFrame(all_containers)
    containers = containers.groupby(
        ['Konteyner_Raqam', 'Mahsulot_Normalized',
         'Transit_Kun', 'Konteyner_Turi', 'Holat', 'Kechikish_Kun'],
        as_index=False
    ).agg({
        'Konteyner_Miqdor': 'sum',
        'Konteyner_Vazn_Kg': lambda s: s.sum(min_count=1),  # hammasi NaN bo'lsa NaN qoladi
        'Chiqish_Sana':     'min',
        'Kelish_Sana':      'min',
    })
    print(f"  ✅ Jami: {len(containers):,} qator, "
          f"{containers['Konteyner_Raqam'].nunique()} ta konteyner")
else:
    containers = pd.DataFrame()
    print("  ⚠️  Konteyner ma'lumoti yo'q")


# 5-bosqich (sotuv trendi tahlili) olib tashlandi — ishlatilmaydi


# ============================================================
# 6. BIRLASHTIRISH
# ============================================================

print("\n6. MA'LUMOTLARNI BIRLASHTIRISH...")

today  = datetime.now()

# 2026-07-09 MUHIM TUZATISH: asos endi Min_Zaxira (to'liq katalog),
# Tarix qoldiq fayli EMAS. Eski kodda asos Tarix fayli edi -- agar tovar
# qoldig'i 0 bo'lib, ombor eksporti (Tarix) uni faylga UMUMAN yozmasa,
# o'sha tovar butun Inventar jadvalidan (demak kamomat/buyurtma
# hisobidan ham) tushib qolardi -- aynan eng kerakli paytda "unutilardi".
# Endi HAR BIR Min_Zaxira'dagi tovar doim natijada bo'ladi; qoldig'i
# Tarix faylida topilmasa 0 deb olinadi.
qoldiq_key = qoldiq.rename(columns={'Mahsulot_Normalized': 'Mahsulot_Key'})

if not min_zaxira.empty:
    # OUTER birlashtirish -- ikkala tomonni ham to'liq saqlaydi:
    #  - Min_Zaxira'da bor, Tarix'da yo'q (qoldiq=0, eksportda ko'rsatilmagan)
    #  - Tarix'da bor, Min_Zaxira'da hali yo'q (katalogga qo'shilmagan tovar)
    # Faqat min_zaxira asosida (how='left') qilinsa 373 ta haqiqiy tovar
    # (masalan "Пр. 30х30 ст 2,0 (304 марка)") yo'qolib qolgani aniqlandi --
    # shu sabab 'outer' tanlandi.
    result = pd.merge(
        min_zaxira.rename(columns={'Mahsulot_Normalized': 'Mahsulot_Key'}),
        qoldiq_key[['Mahsulot_Key', 'Mahsulot', 'Qoldiq_Dona', 'Qoldiq_Summa']],
        on='Mahsulot_Key', how='outer'
    )
else:
    result = qoldiq_key.copy()
    result['Sotuv_Min']      = 0
    result['Kunlik_Istemol'] = 0
    result['Cex_Min']        = 0

for col in ['Sotuv_Min', 'Kunlik_Istemol', 'Cex_Min']:
    result[col] = result[col].fillna(0)

# Konteynerlarni ulash (fast va standart alohida)
if not containers.empty:
    # Kun qoldi hisoblash: kechikkanlar uchun 0 (allaqachon kechikkan)
    containers['Kun_Qoldi'] = containers.apply(
        lambda r: max(0, (pd.to_datetime(r['Kelish_Sana']) - today).days)
        if pd.notna(r['Kelish_Sana']) else None,
        axis=1
    )
    yolda = containers[containers['Holat'] != 'КЕЛДИ ✅'].copy()

    fast_df = (yolda[yolda['Transit_Kun'] == FAST_KUN]
               .groupby('Mahsulot_Normalized')
               .agg(Fast_Miqdor=('Konteyner_Miqdor', 'sum'),
                    Fast_Kelish_Kun=('Kun_Qoldi', 'min'),
                    Fast_Kelish_Sana=('Kelish_Sana', 'min'))
               .reset_index()
               .rename(columns={'Mahsulot_Normalized': 'Mahsulot_Key'}))

    # "12 metrlik" (M12_KUN) — 2026-07-08 qo'shildi. Avval faqat FAST/
    # STANDART ikkita bo'lak bor edi — M12_KUN qatorlar hech biriga mos
    # kelmasdi va butunlay "Йўлда_Жами"dan tushib qolardi (jiddiy xato).
    m12_df = (yolda[yolda['Transit_Kun'] == M12_KUN]
              .groupby('Mahsulot_Normalized')
              .agg(M12_Miqdor=('Konteyner_Miqdor', 'sum'),
                   M12_Kelish_Kun=('Kun_Qoldi', 'min'),
                   M12_Kelish_Sana=('Kelish_Sana', 'min'))
              .reset_index()
              .rename(columns={'Mahsulot_Normalized': 'Mahsulot_Key'}))

    standart_df = (yolda[yolda['Transit_Kun'] == YOLDA_KUN]
                   .groupby('Mahsulot_Normalized')
                   .agg(Standart_Miqdor=('Konteyner_Miqdor', 'sum'),
                        Standart_Kelish_Kun=('Kun_Qoldi', 'min'),
                        Standart_Kelish_Sana=('Kelish_Sana', 'min'))
                   .reset_index()
                   .rename(columns={'Mahsulot_Normalized': 'Mahsulot_Key'}))

    result = (result
              .merge(fast_df,     on='Mahsulot_Key', how='left')
              .merge(m12_df,      on='Mahsulot_Key', how='left')
              .merge(standart_df, on='Mahsulot_Key', how='left'))
else:
    for col in ['Fast_Miqdor', 'Fast_Kelish_Kun', 'Fast_Kelish_Sana',
                'M12_Miqdor', 'M12_Kelish_Kun', 'M12_Kelish_Sana',
                'Standart_Miqdor', 'Standart_Kelish_Kun', 'Standart_Kelish_Sana']:
        result[col] = None

for col in ['Fast_Miqdor', 'M12_Miqdor', 'Standart_Miqdor', 'Qoldiq_Dona', 'Qoldiq_Summa',
            'Sotuv_Min', 'Cex_Min']:
    result[col] = result[col].fillna(0)

result['Mahsulot'] = result['Mahsulot'].fillna(result['Mahsulot_Key'])


# ============================================================
# 7. TAHLIL — SIMULATSIYA VA HOLAT
# ============================================================

print("\n7. TAHLIL...")

result['Yolda_Jami'] = (
    result['Fast_Miqdor'].fillna(0) + result['M12_Miqdor'].fillna(0)
    + result['Standart_Miqdor'].fillna(0)
)
result['Umumiy_Zaxira']      = result['Qoldiq_Dona'] + result['Yolda_Jami']
result['Kategoriya']         = result['Mahsulot'].apply(get_category)
result['Kategoriya_Display'] = result['Mahsulot'].apply(get_category_with_marka)

# Tur belgisi — Цех_Захира > 0 bo'lsa, demak bu tovar Tsexda ishlatiladi.
# E'tibor: bir tovar HAM Sotuv_Min, HAM Cex_Min'ga ega bo'lishi mumkin —
# bunday holda u ЦЕХ🏭 deb belgilanadi, lekin Жами_Мин ikkalasini ham qamrab oladi.
result['Tur'] = result['Cex_Min'].apply(lambda x: 'ЦЕХ🏭' if x > 0 else 'САВДО')

# Жами минимал захира — Холат (КРИТИК/ПАСТ/НОРМА) shu asosda hisoblanadi
result['Min_Zaxira'] = result['Sotuv_Min'] + result['Cex_Min']


def _simulatsiya(row) -> tuple:
    kunlik = float(row.get('Kunlik_Istemol', 0) or 0)
    if kunlik <= 0:
        jami = (float(row.get('Qoldiq_Dona', 0) or 0) +
                float(row.get('Fast_Miqdor', 0) or 0) +
                float(row.get('M12_Miqdor', 0) or 0) +
                float(row.get('Standart_Miqdor', 0) or 0))
        return None, jami, False

    # 2026-07-10 (tuzatildi): Кун >= 0 (avval > 0 edi -- shu sabab BUGUN
    # (Кун=0) keladigan konteyner simulyatsiyadan butunlay TASHLAB
    # YUBORILARDI, xuddi u yo'lda yo'qdek. Natijada masalan Ф-51 ст 0,9
    # (5,8м)(201): Қолдиқ=3186, Йўлда=8003 (Кун=0) bo'lsa ham 8003
    # hisobga olinmay, noto'g'ri КРИТИК chiqardi (to'g'risi: 3186+8003=
    # 11189 >> Мин_Захира=7500 -> НОРМА bo'lishi kerak edi).
    konteynerlar = []
    if (row.get('Fast_Miqdor', 0) > 0
            and pd.notna(row.get('Fast_Kelish_Kun'))
            and (row.get('Fast_Kelish_Kun') or 0) >= 0):
        konteynerlar.append({'kun': float(row['Fast_Kelish_Kun']),
                             'miqdor': float(row['Fast_Miqdor'])})
    if (row.get('M12_Miqdor', 0) > 0
            and pd.notna(row.get('M12_Kelish_Kun'))
            and (row.get('M12_Kelish_Kun') or 0) >= 0):
        konteynerlar.append({'kun': float(row['M12_Kelish_Kun']),
                             'miqdor': float(row['M12_Miqdor'])})
    if (row.get('Standart_Miqdor', 0) > 0
            and pd.notna(row.get('Standart_Kelish_Kun'))
            and (row.get('Standart_Kelish_Kun') or 0) >= 0):
        konteynerlar.append({'kun': float(row['Standart_Kelish_Kun']),
                             'miqdor': float(row['Standart_Miqdor'])})
    konteynerlar.sort(key=lambda x: x['kun'])

    joriy_qoldiq = float(row.get('Qoldiq_Dona', 0) or 0)
    joriy_kun    = 0.0
    tugadi       = False

    for k in konteynerlar:
        sarflanadi    = kunlik * (k['kun'] - joriy_kun)
        joriy_qoldiq -= sarflanadi
        if joriy_qoldiq <= 0:
            tugadi       = True
            joriy_qoldiq = 0
        joriy_qoldiq += k['miqdor']
        joriy_kun     = k['kun']

    kun_yetadi = round(joriy_kun + joriy_qoldiq / kunlik, 0) if kunlik > 0 else None
    return kun_yetadi, round(joriy_qoldiq, 0), tugadi


sim = result.apply(_simulatsiya, axis=1)
result['Kun_Yetadi']                = sim.apply(lambda x: x[0])
result['Yakuniy_Qoldiq']            = sim.apply(lambda x: x[1])
result['Konteyner_Kelguncha_Tugadi']= sim.apply(lambda x: x[2])


def _get_status(row) -> str:
    if row['Min_Zaxira'] == 0:
        return 'МЕЁР ЙЎҚ'
    if row['Konteyner_Kelguncha_Tugadi']:
        return '🔴 КРИТИК'
    yq  = row['Yakuniy_Qoldiq']
    mz  = row['Min_Zaxira']
    if yq <= 0:         return '🔴 КРИТИК'
    elif yq < mz:       return '🔴 КРИТИК'
    elif yq < mz * 1.5: return '🟡 ПАСТ'
    else:               return '🟢 НОРМА'


result['Holat']      = result.apply(_get_status, axis=1)
result['Farq']       = result['Yakuniy_Qoldiq'] - result['Min_Zaxira']
result['Etishmaydi'] = result['Farq'].apply(lambda x: abs(x) if x < 0 else 0)

# Keraksiz tovarlarni olib tashlash
oldin  = len(result)
result = result[~result['Mahsulot'].apply(keraksizmi)].copy()
print(f"  ✅ Keraksiz tovarlar olib tashlandi: {oldin - len(result)} ta")


# ============================================================
# 8. YAKUNIY JADVAL
# ============================================================

final_df = result[[
    'Mahsulot', 'Holat', 'Tur', 'Kategoriya', 'Kategoriya_Display',
    'Qoldiq_Dona',
    'Fast_Miqdor', 'Fast_Kelish_Kun', 'Fast_Kelish_Sana',
    'M12_Miqdor', 'M12_Kelish_Kun', 'M12_Kelish_Sana',
    'Standart_Miqdor', 'Standart_Kelish_Kun', 'Standart_Kelish_Sana',
    'Yolda_Jami', 'Umumiy_Zaxira', 'Min_Zaxira', 'Sotuv_Min', 'Cex_Min',
    'Kunlik_Istemol', 'Kun_Yetadi', 'Yakuniy_Qoldiq',
    'Farq', 'Etishmaydi', 'Qoldiq_Summa',
]].copy()

final_df['Marka']               = result['Mahsulot'].apply(get_marka)
final_df['Kategoriya_Tartib']   = final_df['Kategoriya'].apply(get_category_order)
final_df['Kategoriya_Ierarxiya']= final_df.apply(
    lambda r: r['Kategoriya'] if not r['Marka']
    else f"{r['Kategoriya']} / {r['Marka']}",
    axis=1
)

# Sana ustunlarini formatlash
for sana_col in ['Fast_Kelish_Sana', 'M12_Kelish_Sana', 'Standart_Kelish_Sana']:
    final_df[sana_col] = (pd.to_datetime(final_df[sana_col], errors='coerce')
                          .dt.strftime('%d.%m.%Y')
                          .fillna('—'))

# Ustun nomlarini o'zbek/rus tiliga tarjima
final_df = final_df.rename(columns={
    'Mahsulot':             'Товар',
    'Holat':                'Холат',
    'Tur':                  'Тур',
    'Kategoriya':           'Категория',
    'Kategoriya_Display':   'Категория_Кўриниш',
    'Marka':                'Марка',
    'Qoldiq_Dona':          'Қолдиқ',
    'Fast_Miqdor':          'Тез_Миқдор',
    'Fast_Kelish_Kun':      'Тез_Кун',
    'Fast_Kelish_Sana':     'Тез_Сана',
    'M12_Miqdor':           '12М_Миқдор',
    'M12_Kelish_Kun':       '12М_Кун',
    'M12_Kelish_Sana':      '12М_Сана',
    'Standart_Miqdor':      'Стандарт_Миқдор',
    'Standart_Kelish_Kun':  'Стандарт_Кун',
    'Standart_Kelish_Sana': 'Стандарт_Сана',
    'Yolda_Jami':           'Йўлда_Жами',
    'Umumiy_Zaxira':        'Умумий_Захира',
    'Min_Zaxira':           'Мин_Захира',
    'Sotuv_Min':            'Сотув_Захира',
    'Cex_Min':              'Цех_Захира',
    'Kunlik_Istemol':       'Кунлик_Истеъмол',
    'Kun_Yetadi':           'Кун_Етади',
    'Yakuniy_Qoldiq':       'Якуний_Қолдиқ',
    'Farq':                 'Фарқ',
    'Etishmaydi':           'Етишмайди',
    'Qoldiq_Summa':         'Қолдиқ_Сумма',
    'Kategoriya_Tartib':    'Категория_Тартиб',
    'Kategoriya_Ierarxiya': 'Категория_Иерархия',
})

status_order = {'🔴 КРИТИК': 1, '🟡 ПАСТ': 2, '🟢 НОРМА': 3, 'МЕЁР ЙЎҚ': 4}
final_df['_sort'] = final_df['Холат'].map(status_order).fillna(5)
final_df = final_df.sort_values(['_sort', 'Кун_Етади']).drop(columns=['_sort'])


# ============================================================
# 9. SAQLASH
# ============================================================

print("\n9. NATIJANI SAQLASH...")

# Йўлдаги_Контейнерлар uchun bo'sh jadval ustunlari
YOLDA_COLUMNS = [
    'Контейнер', 'Холат', 'Тури', 'Транзит_Кун',
    'Товар_Тури_Сони', 'Жами_Миқдор', 'Келиш_Санаси', 'Кун_Қолди'
]

with pd.ExcelWriter(str(OUTPUT_FILE), engine='openpyxl') as writer:

    # ── Инвентар — asosiy jadval ────────────────────────────
    final_df.to_excel(writer, sheet_name='Инвентар', index=False)

    # ── Ишлаб_Чиқариш — FAQAT ЦЕХ tovarlari (Цех_Захира > 0) ──
    cex_inv = final_df[final_df['Тур'] == 'ЦЕХ🏭'].copy()
    if not cex_inv.empty:
        cex_inv.to_excel(writer, sheet_name='Ишлаб_Чиқариш', index=False)
        print(f"  ✅ Ишлаб_Чиқариш: {len(cex_inv)} ta tovar")

    # ── Критик ──────────────────────────────────────────────
    kritik_df = final_df[final_df['Холат'] == '🔴 КРИТИК'][[
        'Товар', 'Категория', 'Қолдиқ', 'Мин_Захира', 'Фарқ', 'Кун_Етади'
    ]].copy()
    kritik_df.to_excel(writer, sheet_name='Критик', index=False)


    # ── Контейнерлар ─────────────────────────────────────────
    if not containers.empty:
        cont = containers.copy()
        cont['Кун_Қолди']    = cont['Kelish_Sana'].apply(
            lambda x: max(0, (pd.to_datetime(x) - today).days) if pd.notna(x) else 0
        )
        cont['Категория']    = cont['Mahsulot_Normalized'].apply(get_category)
        cont['Келиш_Санаси'] = cont['Kelish_Sana'].apply(
            lambda x: pd.to_datetime(x, errors='coerce').strftime('%d.%m.%Y')
            if pd.notna(x) else 'НОМАЪЛУМ'
        )
        cont['Юкланган_Сана'] = pd.to_datetime(cont['Chiqish_Sana']).dt.strftime('%d.%m.%Y')
        cont = cont.drop(columns=['Kelish_Sana', 'Chiqish_Sana', 'Kun_Qoldi'],
                         errors='ignore')
        cont = cont.rename(columns={
            'Mahsulot_Normalized': 'Товар',
            'Konteyner_Miqdor':    'Миқдор',
            'Konteyner_Vazn_Kg':   'Вазн_кг',
            'Konteyner_Raqam':     'Контейнер',
            'Konteyner_Turi':      'Тури',
            'Holat':               'Холат',
            'Transit_Kun':         'Транзит_Кун',
            'Kechikish_Kun':       'Кечикиш_Кун',
        })
        cont['Холат_Хисоб'] = cont['Холат'].apply(
            lambda x: 'ЙЎЛДА ЖАМИ 🚢' if str(x) in ('ЙЎЛДА 🚢', 'КЕЧИКДИ ⚠️') else x
        )
        # Kelgan konteynerlar → Юкланган_Сана = '📦 Архив' (srezda alohida ko'rinadi)
        cont.loc[cont['Холат'] == 'КЕЛДИ ✅', 'Юкланган_Сана'] = '📦 Архив'

        # Srez tartibi uchun: eski sana yuqori, yangi past, Архив eng past
        # Sana_Tartib: yoldagilar uchun sana raqami, Архив uchun 99999999
        def _sana_tartib(row):
            if row['Юкланган_Сана'] == '📦 Архив':
                return 99999999
            try:
                qismlar = str(row['Юкланган_Сана']).split('.')
                return int(qismlar[2]) * 10000 + int(qismlar[1]) * 100 + int(qismlar[0])
            except (ValueError, IndexError):
                return 99999998

        cont['Сана_Тартиб'] = cont.apply(_sana_tartib, axis=1)
        cont.sort_values('Келиш_Санаси').to_excel(
            writer, sheet_name='Контейнерлар', index=False
        )

        # Konteyner statistikasi
        keldi    = containers[containers['Holat'] == 'КЕЛДИ ✅']['Konteyner_Raqam'].nunique()
        yolda_n  = containers[containers['Holat'] == 'ЙЎЛДА 🚢']['Konteyner_Raqam'].nunique()
        kechikdi = containers[containers['Holat'].str.startswith('КЕЧИКДИ', na=False)]['Konteyner_Raqam'].nunique()
        yolda_jami = yolda_n + kechikdi

        # ── Контейнер_Хисоб ──────────────────────────────────
        pd.DataFrame([
            {'Холат': 'ЙЎЛДА ЖАМИ 🚢', 'Сон': yolda_jami},
            {'Холат': 'КЕЧИКДИ ⚠️',    'Сон': kechikdi},
            {'Холат': 'КЕЛДИ ✅',       'Сон': keldi},
        ]).to_excel(writer, sheet_name='Контейнер_Хисоб', index=False)

        # ── Контейнер_Статистика ─────────────────────────────
        # Tartib: Йўлда → Кечикди → Келди
        # Тартиб ustuni alohida — card da ko'rinmasin, faqat sort uchun
        stat_df = pd.DataFrame([
            {'Кўрсаткич': 'Йўлда_Жами 🚢',   'Қиймат': yolda_jami},
            {'Кўрсаткич': 'Кечикмоқда ⚠️',   'Қиймат': kechikdi},
            {'Кўрсаткич': 'Келди ✅',          'Қиймат': keldi},
        ])
        stat_df.to_excel(writer, sheet_name='Контейнер_Статистика', index=False)

        # ── Йўлдаги_Контейнерлар ─────────────────────────────
        # MUHIM: Yo'lda konteyner bo'lmasa ham bo'sh varaq yaratiladi
        # Power BI relationship xatosini oldini olish uchun!
        yolda_detail = containers[containers['Holat'] == 'ЙЎЛДА 🚢'].copy()

        if not yolda_detail.empty:
            yolda_detail['Кун_Қолди'] = yolda_detail['Kelish_Sana'].apply(
                lambda x: max(0, (pd.to_datetime(x) - today).days) if pd.notna(x) else 0
            )
            yolda_detail['Келиш_Санаси'] = pd.to_datetime(
                yolda_detail['Kelish_Sana']
            ).dt.strftime('%d.%m.%Y')
            yolda_detail['Категория'] = yolda_detail['Mahsulot_Normalized'].apply(get_category)

            yolda_summary = (yolda_detail
                             .groupby('Konteyner_Raqam')
                             .agg(
                                 Холат=('Holat', 'first'),
                                 Тури=('Konteyner_Turi', 'first'),
                                 Транзит_Кун=('Transit_Kun', 'first'),
                                 Товар_Тури_Сони=('Mahsulot_Normalized', 'nunique'),
                                 Жами_Миқдор=('Konteyner_Miqdor', 'sum'),
                                 Келиш_Санаси=('Келиш_Санаси', 'first'),
                                 Кун_Қолди=('Кун_Қолди', 'min'),
                             )
                             .reset_index()
                             .rename(columns={'Konteyner_Raqam': 'Контейнер'}))
            yolda_summary.sort_values('Кун_Қолди').to_excel(
                writer, sheet_name='Йўлдаги_Контейнерлар', index=False
            )
            print(f"  ✅ Йўлдаги_Контейнерлар: {len(yolda_summary)} ta konteyner")
        else:
            # ── FIX: Yo'lda konteyner yo'q — bo'sh varaq yaratish ──
            # Power BI da "Ключу не соответствует ни одна строка" xatosini
            # oldini olish uchun shu ustunlar bilan bo'sh jadval yoziladi
            pd.DataFrame(columns=YOLDA_COLUMNS).to_excel(
                writer, sheet_name='Йўлдаги_Контейнерлар', index=False
            )
            print(f"  ℹ️  Йўлдаги_Контейнерлар: yo'lda konteyner yo'q (bo'sh varaq yaratildi)")

    else:
        # Konteyner ma'lumoti umuman yo'q — barcha varaqlarni bo'sh yaratish
        pd.DataFrame(columns=['Товар', 'Холат', 'Тури', 'Транзит_Кун',
                               'Миқдор', 'Контейнер', 'Кечикиш_Кун',
                               'Кун_Қолди', 'Категория', 'Келиш_Санаси',
                               'Юкланган_Сана', 'Холат_Хисоб']
                     ).to_excel(writer, sheet_name='Контейнерлар', index=False)

        pd.DataFrame([
            {'Холат': 'ЙЎЛДА ЖАМИ 🚢', 'Сон': 0},
            {'Холат': 'КЕЧИКДИ ⚠️',    'Сон': 0},
            {'Холат': 'КЕЛДИ ✅',       'Сон': 0},
        ]).to_excel(writer, sheet_name='Контейнер_Хисоб', index=False)

        pd.DataFrame([
            {'Кўрсаткич': 'Йўлда_Жами 🚢',   'Қиймат': 0},
            {'Кўрсаткич': 'Кечикмоқда ⚠️',   'Қиймат': 0},
        ]).to_excel(writer, sheet_name='Контейнер_Статистика', index=False)

        pd.DataFrame(columns=YOLDA_COLUMNS).to_excel(
            writer, sheet_name='Йўлдаги_Контейнерлар', index=False
        )
        print(f"  ℹ️  Konteyner yo'q — barcha varaqlar bo'sh yaratildi")

print(f"\n  ✅ Saqlandi: {OUTPUT_FILE}")

# Power BI root fayliga ham nusxa ko'chirish
ROOT_BI_FILE = SCRIPT_DIR / 'NEJAVIYKA_POWER_BI.xlsx'
import shutil as _shutil
_shutil.copy2(str(OUTPUT_FILE), str(ROOT_BI_FILE))
print(f"  ✅ Power BI nusxa: {ROOT_BI_FILE}")


# ============================================================
# 10. YAKUNIY STATISTIKA
# ============================================================

print("\n" + "=" * 100)
print("STATISTIKA")
print("=" * 100)
print(f"  Jami mahsulotlar:    {len(final_df):,}")
print(f"  Cex tovarlari:       {(final_df['Тур'] == 'ЦЕХ🏭').sum():,}")
print(f"  Min zaxira bor:      {(final_df['Мин_Захира'] > 0).sum():,}")
print(f"  Yo'lda mahsulotlar:  {(final_df['Йўлда_Жами'] > 0).sum():,}")
print(f"\n  HOLAT:")
for status, count in final_df['Холат'].value_counts().items():
    print(f"    {status}: {count:,}")

if not containers.empty:
    print(f"\n  KONTEYNER:")
    kechikdi = containers[containers['Holat'].str.startswith('КЕЧИКДИ', na=False)]
    print(f"    ✅ Keldi:    {containers[containers['Holat'] == 'КЕЛДИ ✅']['Konteyner_Raqam'].nunique():>3} ta")
    print(f"    🚢 Yo'lda:   {containers[containers['Holat'] == 'ЙЎЛДА 🚢']['Konteyner_Raqam'].nunique():>3} ta")
    print(f"    ⚠️  Kechikdi: {kechikdi['Konteyner_Raqam'].nunique():>3} ta")
    if not kechikdi.empty:
        print(f"\n    KECHIKKAN KONTEYNERLAR:")
        for _, r in kechikdi.drop_duplicates('Konteyner_Raqam').iterrows():
            chiqish   = pd.to_datetime(r['Chiqish_Sana']).strftime('%d.%m.%Y')
            # DIQQAT (2026-07-08 tuzatildi): avval bu yerda har doim
            # YOLDA_KUN (55) qattiq yozilgan edi — FAST (20) yoki 12M (45)
            # konteyner kechiksa ham noto'g'ri "kelishi kerak edi" sanasi
            # ko'rsatilardi. Endi HAR BIR konteynerning O'Z Transit_Kun
            # qiymati ishlatiladi.
            transit_k = int(r['Transit_Kun']) if pd.notna(r.get('Transit_Kun')) else YOLDA_KUN
            kelish_k  = (pd.to_datetime(r['Chiqish_Sana']) + timedelta(days=transit_k)).strftime('%d.%m.%Y')
            print(f"      • {r['Konteyner_Raqam']} — {chiqish} yuklangan, "
                  f"{kelish_k} kelishi kerak edi, {r['Kechikish_Kun']} kun kechikmoqda")

print("\n" + "=" * 100)
print("✅ TAYYOR!")
print("=" * 100)