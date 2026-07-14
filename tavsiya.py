"""
tavsiya.py — Kunlik sotuv tahlili va Min_Zaxira tavsiyalari
============================================================
Nima qiladi:
  1. tarix/ papkasidan kunlik sotuv hisoblaydi
  2. Min_Zaxira.xlsx bilan taqqoslaydi
  3. Tavsiya.xlsx yaratadi (3 varaq)
  4. Terminalda TOP tovarlarni kategoriya bo'yicha chiqaradi

Ishlatish:
  python tavsiya.py
============================================================
"""

import pandas as pd
import numpy as np
import glob
import os
import math
import re
from datetime import datetime, timedelta
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

from common import (
    normalize_product_name, get_category, get_category_with_marka,
    get_himoya_foiz, yaxlitla_50, hisobla_min_zaxira, min_dan_kunlik_chiqar,
    load_qoldiq_file, keraksizmi, fayl_sanasi,
    YOLDA_KUN, SLIDING_WINDOW,
)

# ============================================================
# YO'LLAR
# ============================================================

SCRIPT_DIR      = Path(__file__).resolve().parent
TARIX_FOLDER    = SCRIPT_DIR / 'tarix'
MIN_ZAXIRA_FILE = SCRIPT_DIR / 'Minimal_zaxiralar' / 'Min_Zaxira.xlsx'
TAVSIYA_FILE    = SCRIPT_DIR / 'chiqish' / 'Tavsiya.xlsx'

# Terminalda har kategoriyadan nechta tovar chiqsin
TOP_N = 10

# Kategoriya tartibi (terminalda shu tartibda chiqadi)
KAT_TARTIB = [
    'ЛИСТ', 'ТРУБА', 'ПРОФИЛЬ', 'БАЛАСИНА', 'СТОЙКА', 'СОККА',
]

# ТРУБА — qattiq tartib (shu nomlar shu tartibda chiqadi)
TRUBA_TARTIB = [
    'Ф-76 ст 0,9', 'Ф-63 ст 0,9', 'Ф-51 ст 0,9',
    'Ф-19 ст 0,9', 'Ф-19 ст 0,7',
    'Ф-16 ст 0,7', 'Ф-16 ст 0,9',
    '(Аркон) Ф-25 ст 0,7', '(Аркон) Ф-25 ст 0,6',
]

# ПРОФИЛЬ — qattiq tartib
PROFIL_TARTIB = [
    'Пр. 50х50 ст 0,9', 'Пр. 40х40 ст 0,9',
    'Пр. 38х38 ст 0,9', 'Пр. 25х25 ст 0,9', 'Пр. 25х25 ст 0,7',
]

# ЛИСТ — qalinlik bo'yicha tartib
LIST_QALINLIK = [0.4, 0.5, 0.6, 0.7, 0.8, 1.0, 1.2, 1.4, 1.45, 1.5, 2.0, 2.5, 3.0, 4.0, 5.0, 6.0, 8.0, 10.0, 12.0]
# Format tartibi: avval kichik (1220), keyin katta (1500)
LIST_FORMAT_TARTIB = ['1220', '1250', '1000', '1500']
# Marka tartibi: 201, 304, 430, 316
LIST_MARKA_TARTIB = ['201', '304', '430', '316', '']

# СОККА — qattiq tartib
SOKKA_TARTIB = ['76', '63', '86', '80', '90', '70', '50']

print("=" * 80)
print("TAVSIYA TIZIMI — Kunlik sotuv tahlili")
print(f"Sana: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
print("=" * 80)


# ============================================================
# 1. MIN_ZAXIRA YUKLASH
# ============================================================

print("\n1. Min_Zaxira yuklanmoqda...")
try:
    mz = pd.read_excel(MIN_ZAXIRA_FILE, sheet_name='Min_Zaxira')
    mz = mz[mz['Товар'].notna()].copy()
    mz['Мин_Захира']      = pd.to_numeric(mz['Мин_Захира'], errors='coerce').fillna(0)
    mz['Кунлик_Истеъмол'] = pd.to_numeric(mz['Кунлик_Истеъмол'], errors='coerce').fillna(0)
    mz['Mahsulot_Normalized'] = mz['Товар'].apply(normalize_product_name)
    print(f"   ✅ {len(mz):,} ta tovar")
except FileNotFoundError:
    print(f"   ❌ Min_Zaxira.xlsx topilmadi!")
    exit(1)


# ============================================================
# 2. TARIX FAYLLARIDAN KUNLIK SOTUV
# ============================================================

print("\n2. Tarix fayllaridan sotuv hisoblanmoqda...")

fayllar = sorted(glob.glob(str(TARIX_FOLDER / '*.xlsx')), key=fayl_sanasi)
if len(fayllar) < 2:
    print("   ❌ Kamida 2 ta tarix fayli kerak!")
    exit(1)

print(f"   📂 {len(fayllar)} ta tarix fayli")

kunlar = []
for f in fayllar:
    fn = os.path.basename(f).replace('.xlsx', '')
    try:
        sana = datetime.strptime(fn, '%d.%m.%Y')
    except ValueError:
        continue
    df_t = load_qoldiq_file(f)
    df_t['Sana'] = sana
    kunlar.append(df_t[['Mahsulot_Normalized', 'Qoldiq_Dona', 'Sana']])

tarix    = pd.concat(kunlar, ignore_index=True).sort_values('Sana')
sanalar  = sorted(tarix['Sana'].unique())
jami_kun = max(1, (sanalar[-1] - sanalar[0]).days)

# Kunlik sotuv hisoblash — faqat kamaygan qoldiq (yuk kelsa hisoblanmaydi)
sotuv_rows = []
for i in range(1, len(sanalar)):
    kecha_s   = sanalar[i-1]
    bugun_s   = sanalar[i]
    kun_farqi = max(1, (bugun_s - kecha_s).days)

    kecha_df = (tarix[tarix['Sana'] == kecha_s]
                [['Mahsulot_Normalized', 'Qoldiq_Dona']]
                .rename(columns={'Qoldiq_Dona': 'Kecha'}))
    bugun_df = (tarix[tarix['Sana'] == bugun_s]
                [['Mahsulot_Normalized', 'Qoldiq_Dona']]
                .rename(columns={'Qoldiq_Dona': 'Bugun'}))

    merged = pd.merge(kecha_df, bugun_df, on='Mahsulot_Normalized', how='inner')
    # Faqat kamaygan = sotuv, oshgan = yuk keldi (o'tkazib yuboriladi)
    merged = merged[merged['Kecha'] > merged['Bugun']].copy()
    merged['Sotuv'] = ((merged['Kecha'] - merged['Bugun']) / kun_farqi).round(3)
    merged['Sana']  = bugun_s
    sotuv_rows.append(merged[['Mahsulot_Normalized', 'Sotuv', 'Sana']])

kunlik_df = pd.concat(sotuv_rows, ignore_index=True)

# Sliding window — oxirgi 30 kun
max_sana     = kunlik_df['Sana'].max()
window_start = max_sana - timedelta(days=SLIDING_WINDOW)
window_df    = kunlik_df[kunlik_df['Sana'] > window_start]

# Jami sotuv / jami kun soni (bir marta katta sotuv ta'sir qilmasin)
jami_sotuv = window_df.groupby('Mahsulot_Normalized').agg(
    Jami_Sotuv=('Sotuv', 'sum'),
    Kuzatuv_Kun=('Sana', 'count')
).reset_index()

# Sliding window kun soni (faqat sotuvda qatnashgan kunlar emas, barcha kunlar)
window_kun_soni = max(1, (max_sana - window_start).days)

# Kunlik real = jami sotuv / barcha kunlar (30 kun)
jami_sotuv['Kunlik_Real'] = (jami_sotuv['Jami_Sotuv'] / window_kun_soni).round(3)

ortacha = jami_sotuv[['Mahsulot_Normalized', 'Kunlik_Real', 'Kuzatuv_Kun']].copy()
ortacha['Kategoriya'] = ortacha['Mahsulot_Normalized'].apply(get_category)

print(f"   ✅ {len(ortacha):,} tovar, {jami_kun} kun tarix")


# ============================================================
# 3. MIN_ZAXIRA BILAN TAQQOSLASH
# ============================================================

print("\n3. Min_Zaxira bilan taqqoslanmoqda...")

tahlil = pd.merge(
    ortacha,
    mz[['Mahsulot_Normalized', 'Мин_Захира', 'Кунлик_Истеъмол']],
    on='Mahsulot_Normalized', how='left'
)
tahlil['Мин_Захира']      = tahlil['Мин_Захира'].fillna(0)
tahlil['Кунлик_Истеъмол'] = tahlil['Кунлик_Истеъмол'].fillna(0)

# Bazaviy kunlik (Min_Zaxira dan teskari hisob)
tahlil['Kunlik_Baza'] = tahlil.apply(
    lambda r: min_dan_kunlik_chiqar(r['Мин_Захира'], r['Kategoriya'])
    if r['Мин_Захира'] > 0 else r['Кунлик_Истеъмол'],
    axis=1
)

# Trend foizi
tahlil['Trend_Foiz'] = tahlil.apply(
    lambda r: round(
        (r['Kunlik_Real'] - r['Kunlik_Baza']) / r['Kunlik_Baza'] * 100, 1
    ) if r['Kunlik_Baza'] > 0.01 else 0,
    axis=1
)

# Holat
def _holat(row):
    foiz = row['Trend_Foiz']
    mz   = row['Мин_Захира']
    real = row['Kunlik_Real']
    kat  = row['Kategoriya']
    if mz == 0:
        return 'МЕЁР ЙЎҚ'
    if foiz > 20:
        tavsiya = yaxlitla_50(real * YOLDA_KUN * (1 + get_himoya_foiz(kat)))
        return f'📈 ОШЯПТИ (+{foiz:.0f}%) → таvsiya: {tavsiya}'
    elif foiz < -20:
        tavsiya = yaxlitla_50(real * YOLDA_KUN * (1 + get_himoya_foiz(kat)))
        return f'📉 КАМАЯПТИ ({foiz:.0f}%) → таvsiya: {tavsiya}'
    else:
        return f'➡️ БАРҚАРОР ({foiz:+.0f}%)'

tahlil['Holat']       = tahlil.apply(_holat, axis=1)
tahlil['Tavsiya_Min'] = tahlil.apply(
    lambda r: yaxlitla_50(r['Kunlik_Real'] * YOLDA_KUN * (1 + get_himoya_foiz(r['Kategoriya'])))
    if r['Мин_Захира'] > 0 else 0,
    axis=1
)
tahlil['Farq_Foiz'] = tahlil['Trend_Foiz']

# Keraksizlarni olib tashlash
tahlil = tahlil[~tahlil['Mahsulot_Normalized'].apply(keraksizmi)].copy()


# ============================================================
# 4. EXCEL SAQLASH
# ============================================================

print("\n4. Tavsiya.xlsx saqlanmoqda...")

oshayotgan  = tahlil[tahlil['Holat'].str.startswith('📈')].sort_values('Trend_Foiz', ascending=False)
kamayyotgan = tahlil[tahlil['Holat'].str.startswith('📉')].sort_values('Trend_Foiz')
barqaror    = tahlil[tahlil['Holat'].str.startswith('➡️')].sort_values('Mahsulot_Normalized')

def _excel_df(df):
    out = df[['Mahsulot_Normalized', 'Kategoriya', 'Kunlik_Real',
              'Кунлик_Истеъмол', 'Мин_Захира', 'Tavsiya_Min',
              'Trend_Foiz', 'Kuzatuv_Kun', 'Holat']].copy()
    out.columns = [
        'Товар', 'Категория', 'Кунлик_Реал',
        'Кунлик_База', 'Хозир_Мин', 'Таvsия_Мин',
        'Тренд_%', 'Кузатув_Кун', 'Холат'
    ]
    out.insert(0, '№', range(1, len(out)+1))
    return out

with pd.ExcelWriter(str(TAVSIYA_FILE), engine='openpyxl') as writer:
    _excel_df(oshayotgan).to_excel(writer,  sheet_name='📈 Ошяпти',   index=False)
    _excel_df(kamayyotgan).to_excel(writer, sheet_name='📉 Камаяпти', index=False)
    _excel_df(barqaror).to_excel(writer,    sheet_name='➡️ Барқарор', index=False)

    # Umumiy varaq
    pd.DataFrame([{
        'Параметр': 'Сана',            'Қиймат': datetime.now().strftime('%d.%m.%Y %H:%M')},
        {'Параметр': 'Жами кун',       'Қиймат': jami_kun},
        {'Параметр': 'Sliding window', 'Қиймат': f'{SLIDING_WINDOW} kun'},
        {'Параметр': 'Ошяпти',         'Қиймат': len(oshayotgan)},
        {'Параметр': 'Камаяпти',       'Қиймат': len(kamayyotgan)},
        {'Параметр': 'Барқарор',       'Қиймат': len(barqaror)},
    ]).to_excel(writer, sheet_name='Маълумот', index=False)

print(f"   ✅ Saqlandi: {TAVSIYA_FILE}")
print(f"      📈 Oshayotgan:   {len(oshayotgan):,} ta")
print(f"      📉 Kamayyotgan:  {len(kamayyotgan):,} ta")
print(f"      ➡️  Barqaror:     {len(barqaror):,} ta")


# ============================================================
# 5. TERMINAL — BELGILANGAN TARTIBDA
# ============================================================

import re as _re

def _chiqar_qator(r):
    nom   = str(r['Mahsulot_Normalized'])[:45]
    real  = r['Kunlik_Real']
    mz    = r['Мин_Захира']
    tav   = r['Tavsiya_Min']
    trend = r['Trend_Foiz']
    if trend > 20:   belgi = '📈'
    elif trend < -20: belgi = '📉'
    else:             belgi = '➡️ '
    print(f"  {nom:<46} {real:>7.1f} {mz:>7.0f} {tav:>8.0f}  {belgi} {trend:+.0f}%")

def _sarlavha(kat, soni):
    print(f"\n{'─'*85}")
    print(f"  {kat}  ({soni} ta tovar)")
    print(f"{'─'*85}")
    print(f"  {'Tovar':<46} {'Real':>7} {'Min':>7} {'Tavsiya':>8}  Holat")
    print(f"  {'-'*80}")

def _past_chiqar(df):
    past = df[df['Trend_Foiz'] < -20].sort_values('Trend_Foiz')
    if not past.empty:
        print(f"  {'─'*80}")
        print(f"  📉 PASAYYOTGANLAR:")
        for _, r in past.iterrows():
            _chiqar_qator(r)

muhim = tahlil[tahlil['Мин_Захира'] > 0].copy()

print("\n" + "=" * 85)
print("HISOBOT — BELGILANGAN TARTIBDA")
print("=" * 85)

# ── ЛИСТ ────────────────────────────────────────────────────
list_df = muhim[muhim['Kategoriya'] == 'ЛИСТ'].copy()
if not list_df.empty:
    _sarlavha('ЛИСТ', len(list_df))

    # Faqat Матовый tovarlar, marka va qalinlik tartibida
    MARKA_TARTIB = {'201': 0, '304': 1, '430': 2, '316': 3}
    FORMAT_TARTIB = {1220: 0, 1250: 1, 1000: 2, 1500: 3}
    QALINLIK_TARTIB = [0.4, 0.5, 0.6, 0.7, 0.8, 1.0, 1.2, 1.4, 1.45, 1.5, 2.0, 2.5, 3.0, 4.0, 5.0, 6.0, 8.0, 10.0, 12.0]

    def _list_sort_key(nom):
        s = str(nom)
        m = _re.search(r'\((\d+)\s*марка\)', s)
        marka = m.group(1) if m else '999'
        mt = MARKA_TARTIB.get(marka, 4)
        fmt_m = _re.search(r'(\d{3,4})х', s)
        fmt = int(fmt_m.group(1)) if fmt_m else 9999
        ft = FORMAT_TARTIB.get(fmt, 5)
        q_m = _re.search(r'Лист-?\s*(\d+[,.]?\d*)', s)
        qalinlik = float(q_m.group(1).replace(',', '.')) if q_m else 99
        qi = QALINLIK_TARTIB.index(qalinlik) if qalinlik in QALINLIK_TARTIB else 99
        return (mt, ft, qi)

    # Faqat Матовый tovarlar
    list_matoviy = list_df[list_df['Mahsulot_Normalized'].str.contains('Матовый', na=False)].copy()
    list_matoviy['_sort'] = list_matoviy['Mahsulot_Normalized'].apply(_list_sort_key)
    list_matoviy = list_matoviy.sort_values('_sort').drop(columns=['_sort'])
    for _, r in list_matoviy.iterrows():
        _chiqar_qator(r)
    _past_chiqar(list_matoviy)

# ── ТРУБА ───────────────────────────────────────────────────
# Faqat shu tovarlar, shu tartibda
TRUBA_FILTER = [
    'Ф-76 ст 0,9', 'Ф-63 ст 0,9', 'Ф-51 ст 0,9',
    'Ф-19 ст 0,9', 'Ф-19 ст 0,7',
    'Ф-16 ст 0,7', 'Ф-16 ст 0,9',
    '(Аркон) Ф-25 ст 0,7', '(Аркон) Ф-25 ст 0,6',
]
truba_df = muhim[muhim['Kategoriya'] == 'ТРУБА'].copy()
if not truba_df.empty:
    def _truba_key(nom):
        s = str(nom)
        for i, pat in enumerate(TRUBA_FILTER):
            if pat.lower() in s.lower():
                uzunlik_m = _re.search(r'\((\d+[,.]\d+)\s*м\)', s)
                uzunlik = float(uzunlik_m.group(1).replace(',', '.')) if uzunlik_m else 6.0
                return (i, uzunlik)
        return (999, 0)

    # Faqat belgilangan tovarlar
    truba_fil = truba_df[truba_df['Mahsulot_Normalized'].apply(
        lambda n: any(p.lower() in n.lower() for p in TRUBA_FILTER)
    )].copy()
    truba_fil['_sort'] = truba_fil['Mahsulot_Normalized'].apply(_truba_key)
    truba_fil = truba_fil.sort_values('_sort').drop(columns=['_sort'])
    _sarlavha('ТРУБА', len(truba_fil))
    for _, r in truba_fil.iterrows():
        _chiqar_qator(r)
    _past_chiqar(truba_fil)

# ── ПРОФИЛЬ ─────────────────────────────────────────────────
PROFIL_FILTER = [
    'Пр. 50х50 ст 0,9', 'Пр. 40х40 ст 0,9',
    'Пр. 38х38 ст 0,9', 'Пр. 25х25 ст 0,9', 'Пр. 25х25 ст 0,7',
]
profil_df = muhim[muhim['Kategoriya'] == 'ПРОФИЛЬ'].copy()
if not profil_df.empty:
    def _profil_key(nom):
        s = str(nom)
        for i, pat in enumerate(PROFIL_FILTER):
            if pat.lower() in s.lower():
                uzunlik_m = _re.search(r'\((\d+[,.]\d+)\s*м\)', s)
                uzunlik = float(uzunlik_m.group(1).replace(',', '.')) if uzunlik_m else 6.0
                return (i, uzunlik)
        return (999, 0)

    profil_fil = profil_df[profil_df['Mahsulot_Normalized'].apply(
        lambda n: any(p.lower() in n.lower() for p in PROFIL_FILTER)
    )].copy()
    profil_fil['_sort'] = profil_fil['Mahsulot_Normalized'].apply(_profil_key)
    profil_fil = profil_fil.sort_values('_sort').drop(columns=['_sort'])
    _sarlavha('ПРОФИЛЬ', len(profil_fil))
    for _, r in profil_fil.iterrows():
        _chiqar_qator(r)
    _past_chiqar(profil_fil)

# ── БАЛАСИНА — top 10 ────────────────────────────────────────
bal_df = muhim[muhim['Kategoriya'] == 'БАЛАСИНА'].copy()
if not bal_df.empty:
    _sarlavha('БАЛАСИНА', len(bal_df))
    for _, r in bal_df.nlargest(TOP_N, 'Kunlik_Real').iterrows():
        _chiqar_qator(r)
    _past_chiqar(bal_df)

# ── СТОЙКА — top 10 ─────────────────────────────────────────
stoyka_df = muhim[muhim['Kategoriya'] == 'СТОЙКА'].copy()
if not stoyka_df.empty:
    _sarlavha('СТОЙКА', len(stoyka_df))
    for _, r in stoyka_df.nlargest(TOP_N, 'Kunlik_Real').iterrows():
        _chiqar_qator(r)
    _past_chiqar(stoyka_df)

# ── СОККА — faqat 76, 63, 86 ────────────────────────────────
SOKKA_FILTER = ['76', '63', '86']
sokka_df = muhim[muhim['Kategoriya'] == 'СОККА'].copy()
if not sokka_df.empty:
    def _sokka_key(nom):
        s = str(nom)
        for i, d in enumerate(SOKKA_FILTER):
            if f'Сокка {d}' in s or f'Сокка. Голд {d}' in s:
                gold = 1 if 'Голд' in s else 0
                uzunlik_m = _re.search(r'х(\d+)', s)
                uzunlik = int(uzunlik_m.group(1)) if uzunlik_m else 99
                return (i, gold, uzunlik)
        return (999, 0, 0)

    sokka_fil = sokka_df[sokka_df['Mahsulot_Normalized'].apply(
        lambda n: any(
            f'Сокка {d}' in n or f'Сокка. Голд {d}' in n
            for d in SOKKA_FILTER
        )
    )].copy()
    sokka_fil['_sort'] = sokka_fil['Mahsulot_Normalized'].apply(_sokka_key)
    sokka_fil = sokka_fil.sort_values('_sort').drop(columns=['_sort'])
    _sarlavha('СОККА', len(sokka_fil))
    for _, r in sokka_fil.iterrows():
        _chiqar_qator(r)
    _past_chiqar(sokka_fil)

print(f"\n{'='*85}")
print(f"✅ TAYYOR! Batafsil: {TAVSIYA_FILE}")
print(f"{'='*85}")
