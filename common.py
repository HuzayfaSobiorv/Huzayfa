"""
common.py — Umumiy yordamchi modul
====================================
Bu fayl NEJAVIYKA_v3.py va min_cache_tizimi.py tomonidan import qilinadi.
Normalize, kategoriya, himoya foizi kabi BARCHA umumiy mantiq shu yerda.

MUHIM: Ikki faylda bir xil normalize ishlatilishi shart, aks holda
       cache kalitlari qoldiq kalitlariga mos kelmaydi.
====================================
"""

import re
import math
import os
import pandas as pd
from datetime import datetime
import warnings

warnings.filterwarnings('ignore')

# ============================================================
# KONSTANTALAR — faqat shu yerda, boshqa joyda takrorlanmaydi
# ============================================================

KELISH_KUNI     = 55    # ◄ ASOSIY: konteyner yetib kelish vaqti (kun) — FAQAT SHU YERDA O'ZGARTIR
YOLDA_KUN       = KELISH_KUNI   # alias (eski kod uchun)
FAST_KUN        = 20    # tezkor konteyner yetib kelish vaqti (kun)
TREND_KUN       = KELISH_KUNI   # necha kun barqaror bo'lsa min oshiriladi
SLIDING_WINDOW  = 30    # oxirgi necha kunni kuzatish
SEZON_OYNA      = 7     # sezon taqqoslash oynasi (kun)
MIN_KUZATUV_KUN = 7     # yangi tovar uchun kamida necha kun kuzatish kerak

# Kategoriya bo'yicha himoya yostig'i (buffer foizi)
HIMOYA_FOIZ: dict[str, float] = {
    'ТРУБА':        0.50,
    'ПРОФИЛЬ':      0.50,
    'ЛИСТ':         0.15,
    'ЛИСТ РУЛОН':   0.15,
}
HIMOYA_DEFAULT = 0.50

# Kategoriya tartib raqami (Power BI da saralash uchun)
KATEGORIYA_TARTIB: dict[str, int] = {
    'ТРУБА (201)':    1,  'ТРУБА (304)':    2,  'ТРУБА (430)':    3,
    'ТРУБА (316)':    4,  'ТРУБА (Бесшовный)': 5, 'ТРУБА':         6,
    'ПРОФИЛЬ (201)':  7,  'ПРОФИЛЬ (304)':  8,  'ПРОФИЛЬ (430)':  9,
    'ПРОФИЛЬ (316)': 10,  'ПРОФИЛЬ':       11,
    'ЛИСТ (201)':    12,  'ЛИСТ (304)':    13,  'ЛИСТ (430)':    14,
    'ЛИСТ (316)':    15,  'ЛИСТ':          16,
    'БАЛАСИНА':      17,  'СТОЙКА':        18,  'ЧАШКА':         19,
    'ОТВОД':         20,  'СОККА':         21,  'ШАР':           22,
    'СОВУН':         23,  'БОШҚА':         24,  'ЛИСТ РУЛОН':    99,
}

# Hisobot uchun filtrlanadigan kalit so'zlar
KERAKSIZ_QISM = [
    'тумба', 'ХС №', 'заказ', 'Гардероб', 'консол',
    'ясаб', 'учун', 'дар', 'берилди', 'вешалка', 'ди',
    'стеллаж', 'ЛС-', 'забор', 'Ошхона', 'подставка',
    'Сори', 'степлер', 'кронштейн', 'Стул', 'Сушилка',
    'Табуретка', 'Ю/С', 'гич', 'ЛК-', 'Казиро', 'Германтин', 'тувак', 'Декоративни',
    '100х100-Чашка-02 (Голд)',
    '100х100-Чашка-01 (Голд)',
    '100х100-Чашка-01 (Глянцевый)',
    '100х100-Чашка-02 (Глянцевый)',
    '100х100-Чашка-01 (Матовый)',
    '100х100-Чашка-02 (Матовый)',
    'Декоративний', 'Детский качале', 'Стелаж', 'Умвалик', 'Баласина (Балик)', 'Брак баласина',
    'Фланс','S-гул ','БП','Дрель','Дуга (Сотув булми)','Зажим ', 'Заклёпка','Итого','КР-001 (210х100)',
    'Кабр','Казерог ','Каска','Катта гул','Кислород балон','Контейнер','Краска. Голд','Кресло', 
    'Кук материал','Кучага хизмат сварка хизмати','Ламинад Д=1700','Латок','Ловия',
    'Малочни материал','Масжид','Маска','(Сотув булими)',
    'Мойка 2600х80х40 (Сотув булмига)','Мешалка','Ножка 900Х650Х45 (Сотув булмига) Голд','Нукус сахна','сахна','Обувница (карши)Голд',
    'карши','бухоро','Ойна (кора)','Петля',' (марям)','урожайнига','Тахта','Очки №1','Перфоратор',
    'Перегаротка 210х50 (Ташкент) Голд','Перегароткa',
    'Петля','Перчатка','Тележка','Фанер',
]

KERAKSIZ_ALON = ['стол']   # to'liq so'z sifatida qidiriladi

# Har doim o'tsin — filtrlanmasin
KERAKSIZ_ISTISNO = [
    'стул ойокчаси №04',
    'стул ойокчаси №05',
]


# ============================================================
# NORMALIZE — bu funksiya ikkala faylda bir xil ishlatiladi
# ============================================================

def normalize_product_name(name: str) -> str:
    if pd.isna(name):
        return name

    name = str(name).strip()
    name = ' '.join(name.split())

    def _add_m(match):
        content = match.group(1)
        if 'м' in content:
            return match.group(0)
        if any(c.isdigit() for c in content):
            return f"({content} м)"
        return match.group(0)

    name = re.sub(r'\(([0-9,\.]+)\)(?!\s*марка)', _add_m, name)
    name = re.sub(r'м\s+\)', 'м)', name)
    name = re.sub(r'Лист-(\d)', r'Лист- \1', name)
    name = ' '.join(name.split())
    return name


# ============================================================
# KATEGORIYA ANIQLASH
# ============================================================

def get_category(name) -> str:
    if pd.isna(name):
        return 'БОШҚА'

    s  = str(name).strip()
    sl = s.lower()

    if re.match(r'^(\([^)]*\)\s*)?Ф-\d+', s) and 'ст' in s and re.search(r'\(\d+.*м\)', s):
        return 'ТРУБА'
    if re.match(r'^.*Пр\.\s+\d+х\d+', s) and 'ст' in s and re.search(r'\(\d+.*м\)', s):
        return 'ПРОФИЛЬ'
    if re.match(r'^Лист-\s*\d+', s) and re.search(r'\(\d+х\d+\)', s):
        return 'ЛИСТ'
    if re.match(r'^Лист\s+рулон', s):
        return 'ЛИСТ РУЛОН'

    keyword_map = [
        ('баласина', 'БАЛАСИНА'),
        ('стойка',   'СТОЙКА'),
        ('сокка',    'СОККА'),
        ('шар',      'ШАР'),
        ('отвод',    'ОТВОД'),
        ('кузикорин','КУЗИКОРИН'),
        ('чашка',    'ЧАШКА'),
        ('совун',    'СОВУН'),
    ]
    for keyword, category in keyword_map:
        if keyword in sl:
            return category

    return 'БОШҚА'


def get_marka(name) -> str:
    if pd.isna(name):
        return ''
    match = re.search(r'\((\d+)\s*марка\)', str(name))
    return match.group(1) if match else ''


def get_category_with_marka(name) -> str:
    kat   = get_category(name)
    marka = get_marka(name)
    if marka and kat in ('ТРУБА', 'ПРОФИЛЬ', 'ЛИСТ', 'ЛИСТ РУЛОН'):
        return f"{kat} ({marka})"
    return kat


def get_category_order(kategoriya: str) -> int:
    return KATEGORIYA_TARTIB.get(kategoriya, 99)


# ============================================================
# HIMOYA FOIZI VA MIN HISOBLASH
# ============================================================

def get_himoya_foiz(kategoriya: str) -> float:
    return HIMOYA_FOIZ.get(kategoriya, HIMOYA_DEFAULT)


def yaxlitla_50(qiymat: float) -> int:
    if qiymat <= 0:
        return 0
    return math.ceil(qiymat / 50) * 50


def hisobla_min_zaxira(kunlik_istemol: float, kategoriya: str) -> int:
    foiz = get_himoya_foiz(kategoriya)
    return yaxlitla_50(kunlik_istemol * YOLDA_KUN * (1 + foiz))


def min_dan_kunlik_chiqar(min_zaxira: float, kategoriya: str) -> float:
    foiz = get_himoya_foiz(kategoriya)
    denom = YOLDA_KUN * (1 + foiz)
    if denom <= 0:
        return 0.0
    return round(min_zaxira / denom, 3)


# ============================================================
# QOLDIQ FAYLINI YUKLASH
# ============================================================

def parse_qoldiq_str(text) -> int:
    if pd.isna(text):
        return 0
    parts = str(text).strip().split('/')
    raw = parts[0].strip().replace(',', '').replace(' ', '')
    try:
        return int(raw) if raw and raw != '-' else 0
    except ValueError:
        return 0


def load_qoldiq_file(filepath: str) -> pd.DataFrame:
    df = pd.read_excel(filepath, header=4)
    col_count = len(df.columns)
    mid_cols = [f'_Col{i}' for i in range(1, col_count - 2)]
    df.columns = ['Mahsulot'] + mid_cols + ['Qoldiq_Str', 'Qoldiq_Summa']
    df = df[df['Mahsulot'].notna()].copy()
    df = df[df['Mahsulot'] != 'Товар'].copy()
    df['Qoldiq_Dona']         = df['Qoldiq_Str'].apply(parse_qoldiq_str)
    df['Qoldiq_Summa']        = pd.to_numeric(df['Qoldiq_Summa'], errors='coerce').fillna(0)
    df['Mahsulot_Normalized'] = df['Mahsulot'].apply(normalize_product_name)
    return df[['Mahsulot', 'Mahsulot_Normalized', 'Qoldiq_Dona', 'Qoldiq_Summa']].copy()


# ============================================================
# FILTR — keraksiz tovarlar
# ============================================================

def keraksizmi(nom) -> bool:
    if pd.isna(nom):
        return False
    nom_lower = str(nom).lower()

    # Istisno — bu nomlar har doim o'tsin, filtrlanmasin
    if any(i in nom_lower for i in KERAKSIZ_ISTISNO):
        return False

    # Qism sifatida kirgan bo'lsa
    for qism in KERAKSIZ_QISM:
        if qism.lower() in nom_lower:
            return True

    # To'liq so'z sifatida
    for alon in KERAKSIZ_ALON:
        if re.search(r'\b' + re.escape(alon.lower()) + r'\b', nom_lower):
            return True

    return False


# ============================================================
# FAYL SANA YORDAMCHISI
# ============================================================

def fayl_sanasi(fayl_yoli: str) -> datetime:
    filename = os.path.basename(fayl_yoli).replace('.xlsx', '')
    try:
        return datetime.strptime(filename, '%d.%m.%Y')
    except ValueError:
        return datetime.fromtimestamp(os.path.getmtime(fayl_yoli))