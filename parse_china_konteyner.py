"""
parse_china_konteyner.py — Xitoy multi-konteyner parser (FINAL)
================================================================
Qoidalar:
  - Truba:   qalinlik*diametr*uzunlik_mm
  - Profil:  qalinlik*en*boy*uzunlik_mm  (ba'zida boy < 5 → ×10)
  - List:    qalinlik*en*boy (mm)
  - Qalinlik: doim yuqoriga yaxlitlash
  - 1.35 → truba/profil: avval 1,35 qidir → topilmasa 1,4
  - O'lcham: eng yaqin 10 ga yaxlitlash (1219→1220, 2438→2440)
  - Rang: 砂板→Матовый, 8K→Глянцевый, 8K钛金→Голд, 8K黑钛→Кора
  - Marka: 材质 ustunidan (201/J1 → 201 + Ж-1 suffix)
  - Brend: (Аркон),(Голд),(Кора) — saqlanadi
  - Birlashtirish: raqam mos + sana ±2 kun → bitta fayl
  - Noma'lum: crash yo'q → terminal + NOMALUM.xlsx

FIX (sana muammosi):
  _parse_truba_profil da global_date endi loop boshlanishidan OLDIN
  butun fayl bo'ylab skanlanadi. Shuning uchun sana 柜 sarlavhasida
  emas, yuqori burchakda (02.04.2026 kabi) bo'lsa ham to'g'ri olinadi.
  Oldin: birinchi 柜 sarlavhasida sana yo'q bo'lsa oldingi konteynerning
         sanasi (global_date) qo'llanilar edi → noto'g'ri sana chiqardi.
================================================================
"""

import pandas as pd
import numpy as np
import re
import os
import math
from pathlib import Path
from datetime import datetime, timedelta

# ── Ranglar mapping ──────────────────────────────────────────────────────────
RANG_MAP = {
    '砂板':   'Матовый',
    '8K':     'Глянцевый',
    '8K钛金': 'Голд',
    '8K黑钛': 'Кора',
}

# ── Qalinlik yaxlitlash ──────────────────────────────────────────────────────

def _yaxlitla_qalinlik(xitoy_val: float, tur: str, mavjud_nomlar: set) -> str:
    """
    Xitoy qalinligini bizniki formatiga aylantiradi.

    Qoidalar:
      1. Butun son (1.0, 2.0, 3.0) → o'zini ol: '1,0', '2,0', '3,0'
      2. Kasr son → YUQORIGA yaxlitla (0.85→0,9, 1.45→1,5, 1.95→2,0)
      3. 1.35 MAXSUS (faqat truba/profil):
           mavjud nomlarda 'ст 1,35' bor → '1,35'
           yo'q → '1,4'

    Qaytaradi: '0,9', '1,0', '1,35', '1,4' kabi string
    """
    val = round(xitoy_val, 3)

    val_x10 = round(val * 10)
    if abs(val * 10 - val_x10) < 0.01:
        if val_x10 % 10 == 0:
            return f"{int(val_x10 // 10)},0"
        else:
            s = f"{val:.1f}".replace('.', ',')
            return s

    if abs(val - 1.35) < 0.01:
        if tur in ('truba', 'profil'):
            has_135 = any('ст 1,35' in n for n in mavjud_nomlar)
            return '1,35' if has_135 else '1,4'
        else:
            return '1,4'

    yaxlit = math.ceil(round(val * 10, 4)) / 10
    s = f"{yaxlit:.1f}".replace('.', ',')
    return s


# ── O'lcham yaxlitlash (1219→1220) ──────────────────────────────────────────

def _yaxlitla_olcham(val: float) -> int:
    """Eng yaqin 10 ga yuqoriga yaxlitlash: 1219→1220, 2438→2440"""
    return math.ceil(val / 10) * 10


def _list_olchamlar(en_mm: float, boy_mm: float) -> tuple:
    if 2600 <= boy_mm <= 2700:
        return 1250, 2700
    return _yaxlitla_olcham(en_mm), _yaxlitla_olcham(boy_mm)


# ── Rang aniqlash ────────────────────────────────────────────────────────────

def _rang_topish(rang_str: str) -> str:
    if not rang_str or pd.isna(rang_str):
        return ''
    s = str(rang_str).strip()
    if s in RANG_MAP:
        return RANG_MAP[s]
    for key, val in sorted(RANG_MAP.items(), key=lambda x: -len(x[0])):
        if key in s:
            return val
    return s


# ── Marka va suffix ajratish ─────────────────────────────────────────────────

def _marka_ajrat(marka_str: str) -> tuple[str, str]:
    if not marka_str or pd.isna(marka_str):
        return '', ''
    s = str(marka_str).strip()
    m = re.match(r'^(\d+)(?:/J(\d+))?$', s)
    if m:
        marka = m.group(1)
        suffix = f'(Ж-{m.group(2)})' if m.group(2) else ''
        return marka, suffix
    raqam = re.search(r'\d+', s)
    return raqam.group() if raqam else s, ''


# ── Truba nomi yaratish ──────────────────────────────────────────────────────

def _truba_nomi(parts: list, marka: str, suffix: str,
                brend: str, mavjud_nomlar: set) -> str | None:
    if len(parts) != 3:
        return None
    try:
        qalinlik_x  = float(parts[0])
        diametr     = float(parts[1])
        uzunlik_mm  = float(parts[2])
    except ValueError:
        return None

    qalinlik_s  = _yaxlitla_qalinlik(qalinlik_x, 'truba', mavjud_nomlar)
    diametr_int = int(round(diametr))
    uzunlik_m   = uzunlik_mm / 1000
    if uzunlik_m == int(uzunlik_m):
        uzunlik_s = str(int(uzunlik_m))
    else:
        uzunlik_s = f"{uzunlik_m:.1f}".replace('.', ',').rstrip('0').rstrip(',')
        if ',' not in uzunlik_s:
            uzunlik_s = uzunlik_s

    nom = f"Ф-{diametr_int} ст {qalinlik_s}"
    if brend == 'Бесшовный':
        nom += f" Бесшовный ({uzunlik_s} м)"
    else:
        nom += f" ({uzunlik_s} м)"
        if brend:
            nom = f"({brend}) {nom}"
    if marka:
        nom += f" ({marka} марка)"
    if suffix:
        nom += f" {suffix}"
    return nom


# ── Profil nomi yaratish ─────────────────────────────────────────────────────

def _profil_nomi(parts: list, marka: str, suffix: str,
                 brend: str, mavjud_nomlar: set) -> str | None:
    if len(parts) != 4:
        return None
    try:
        qalinlik_x = float(parts[0])
        en         = float(parts[1])
        boy        = float(parts[2])
        uzunlik_mm = float(parts[3])
    except ValueError:
        return None

    if boy < 5:
        boy = boy * 10

    qalinlik_s = _yaxlitla_qalinlik(qalinlik_x, 'profil', mavjud_nomlar)
    en_int     = int(round(en))
    boy_int    = int(round(boy))
    uzunlik_m  = uzunlik_mm / 1000

    if uzunlik_m == int(uzunlik_m):
        uzunlik_s = str(int(uzunlik_m))
    else:
        uzunlik_s = f"{uzunlik_m:.1f}".replace('.', ',')

    nom = f"Пр. {en_int}х{boy_int} ст {qalinlik_s} ({uzunlik_s} м)"
    if brend:
        nom = f"({brend}) {nom}"
    if marka:
        nom += f" ({marka} марка)"
    if suffix:
        nom += f" {suffix}"
    return nom


# ── List nomi yaratish ───────────────────────────────────────────────────────

def _list_nomi(parts: list, marka: str, rang: str) -> str | None:
    if len(parts) != 3:
        return None
    try:
        qalinlik_x = float(parts[0])
        en_mm      = float(parts[1])
        boy_mm     = float(parts[2])
    except ValueError:
        return None

    qalinlik_s = _yaxlitla_qalinlik(qalinlik_x, 'list', set())
    en_int, boy_int = _list_olchamlar(en_mm, boy_mm)

    nom = f"Лист-{qalinlik_s} ({en_int}х{boy_int})"
    if rang:
        nom += f" ({rang})"
    if marka:
        nom += f" ({marka} марка)"
    return nom


# ── Spec → nom (asosiy konvertatsiya) ────────────────────────────────────────

def spec_to_nom(spec: str, marka_raw: str, rang_raw: str,
                mavjud_nomlar: set, brend_override: str = '') -> tuple[str, bool]:
    if not spec or pd.isna(spec):
        return str(spec), False

    s = str(spec).strip()

    # Brend prefiksini specdan ajratish: '(Аркон)0.65*25*6000' → brend='Аркон', spec='0.65*25*6000'
    BRENDLAR = ['Аркон', 'Голд', 'Кора', 'Бесшовный']
    brend_from_spec = ''
    for b in BRENDLAR:
        if s.startswith(f'({b})'):
            brend_from_spec = b
            s = s[len(f'({b})'):]
            break
        elif s.lower().startswith(b.lower()):
            brend_from_spec = b
            s = s[len(b):].strip()
            break

    parts = [p.strip() for p in s.split('*')]

    marka, suffix = _marka_ajrat(marka_raw)
    rang  = _rang_topish(rang_raw)
    # brend_override ustunlik qiladi, yo'q bo'lsa specdan topilgan brendni ishlatamiz
    brend = brend_override if brend_override else brend_from_spec

    if len(parts) == 3:
        try:
            vals = [float(p) for p in parts]
            if vals[1] > 500 and vals[2] > 500:
                nom = _list_nomi(parts, marka, rang)
            else:
                nom = _truba_nomi(parts, marka, suffix, brend, mavjud_nomlar)
        except ValueError:
            nom = None

    elif len(parts) == 4:
        nom = _profil_nomi(parts, marka, suffix, brend, mavjud_nomlar)

    else:
        nom = None

    if nom:
        return nom, True
    return f"НОМАЪЛУМ: {s}", False


# ── Fayl turini aniqlash (truba/profil vs list) ──────────────────────────────

def _fayl_turi(df_raw: pd.DataFrame) -> str:
    col_texts = ' '.join(str(c) for c in df_raw.values.flatten() if pd.notna(c))
    if '颜色名称' in col_texts or '砂板' in col_texts or '8K' in col_texts:
        return 'list'
    return 'truba_profil'


def _tekshir_truba_fayl(df_raw: pd.DataFrame, fayl_nomi: str) -> list[str]:
    xatolar = []
    col_texts = ' '.join(str(c) for c in df_raw.values.flatten() if pd.notna(c))
    rows = df_raw.values.tolist()

    gui_topildi = False
    for row in rows:
        for cell in row:
            if pd.notna(cell) and re.match(r'^[大小]?柜\d*$', str(cell).strip()):
                gui_topildi = True
                break
        if gui_topildi:
            break
    if not gui_topildi:
        xatolar.append("柜N / 小柜N / 大柜N sarlavhasi topilmadi (konteyner boshi yo'q)")

    if '规格' not in col_texts:
        xatolar.append("'规格' (spec) ustuni topilmadi")
    if '支数' not in col_texts and '件数' not in col_texts:
        xatolar.append("'支数' yoki '件数' (miqdor) ustuni topilmadi")

    raqam_topildi = bool(re.search(r'柜号[：:]\s*[A-Z]{4}\d{7,}', col_texts))
    if not raqam_topildi:
        xatolar.append("'备注' da konteyner raqami topilmadi (柜号:XXXX ko'rinishida bo'lishi kerak)")

    sana_topildi = False
    for row in rows:
        for cell in row:
            if pd.notna(cell):
                if hasattr(cell, 'year'):
                    sana_topildi = True
                    break
                cell_str = re.sub(r'\s+', '', str(cell))
                if re.search(r'\d{2}[./]\d{2}[./]\d{4}', cell_str):
                    sana_topildi = True
                    break
        if sana_topildi:
            break
    if not sana_topildi:
        xatolar.append("Yuklangan sana topilmadi (DD.MM.YYYY ko'rinishida bo'lishi kerak)")

    return xatolar


def _tekshir_list_fayl(df_raw: pd.DataFrame, fayl_nomi: str) -> list[str]:
    xatolar = []
    col_texts = ' '.join(str(c) for c in df_raw.values.flatten() if pd.notna(c))
    rows = df_raw.values.tolist()

    if '颜色名称' not in col_texts:
        xatolar.append("'颜色名称' (rang nomi) ustuni topilmadi")
    if '规格' not in col_texts:
        xatolar.append("'规格' (spec) ustuni topilmadi")
    if '数量' not in col_texts:
        xatolar.append("'数量' (miqdor) ustuni topilmadi")

    raqam_topildi = bool(re.search(r'柜号[：:]\s*[A-Z]{4}\d{7,}', col_texts))
    if not raqam_topildi:
        xatolar.append("'备注' da konteyner raqami topilmadi (柜号:XXXX ko'rinishida bo'lishi kerak)")

    sana_topildi = False
    for row in rows:
        for cell in row:
            if pd.notna(cell):
                if hasattr(cell, 'year'):
                    sana_topildi = True
                    break
                cell_str = re.sub(r'\s+', '', str(cell))
                if re.search(r'\d{2}[./]\d{2}[./]\d{4}', cell_str):
                    sana_topildi = True
                    break
                if re.search(r'\d{1,2}月\d{1,2}日', str(cell)):
                    sana_topildi = True
                    break
        if sana_topildi:
            break
    if not sana_topildi:
        xatolar.append("Yuklangan sana topilmadi")

    return xatolar


# ── Yordamchi: butun fayl bo'ylab birinchi sanani topish ─────────────────────

def _fayldan_sana_topish(rows: list) -> datetime | None:
    """
    Fayl qatorlarini BOSHIDAN ko'rib chiqadi va birinchi to'g'ri sanani qaytaradi.

    Nima uchun kerak:
      Xitoy faylida sana ba'zan 柜 sarlavhasida emas, yuqori burchakda
      (masalan 02.04.2026) bo'ladi. Eski kodda bu sana o'tkazib yuborilardi
      va oldingi fayldan qolgan sana (global_date) ishlatilardi.
      Natijada TCNU8902080 uchun 02.04 o'rniga 30.03 chiqardi.
    """
    for row in rows:
        for cell in row:
            if pd.notna(cell):
                # Excel datetime obyekti
                if hasattr(cell, 'year'):
                    try:
                        return datetime(cell.year, cell.month, cell.day)
                    except Exception:
                        pass
                # String: '02.04.2026' yoki '02/04/2026'
                cell_str = re.sub(r'\s+', '', str(cell))
                m = re.search(r'(\d{2})[./](\d{2})[./](\d{4})', cell_str)
                if m:
                    try:
                        return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
                    except ValueError:
                        pass
    return None


# ── Truba/Profil faylini parse qilish ───────────────────────────────────────

def _parse_truba_profil(df_raw: pd.DataFrame, mavjud_nomlar: set) -> list[dict]:
    """
    Truba/profil Excel faylidan konteynerlar ro'yxatini qaytaradi.

    Format:
      - 柜N sarlavhasi → yangi konteyner
      - 备注 qatorida 柜号:XXXX
      - 规格 ustunida spec, 件数 ustunida miqdor, 材质 ustunida marka
    """
    rows   = df_raw.values.tolist()
    n_rows = len(rows)

    # Ustun indekslarini aniqlash
    spec_col  = None
    dona_col  = None
    marka_col = None

    for ri in range(min(30, n_rows)):
        for ci, cell in enumerate(rows[ri]):
            if pd.notna(cell):
                s = str(cell).strip()
                if s == '规格'  and spec_col  is None: spec_col  = ci
                if s == '支数'  and dona_col  is None: dona_col  = ci  # dona soni (支数)
                if s == '材质'  and marka_col is None: marka_col = ci

    # 支数 topilmasa 件数 ga fallback (eski fayllar uchun)
    if dona_col is None:
        for ri in range(min(30, n_rows)):
            for ci, cell in enumerate(rows[ri]):
                if pd.notna(cell) and str(cell).strip() == '件数':
                    dona_col = ci
                    print(f"   ⚠️  '支数' topilmadi — '件数' (pochka) ishlatilmoqda")
                    break
            if dona_col is not None:
                break

    if spec_col is None or dona_col is None:
        raise ValueError(f"规格 yoki 支数/件数 ustuni topilmadi!")

    # ═══════════════════════════════════════════════════════════════
    # FIX: global_date ni fayl BOSHIDAN o'qish
    # Muammo: 柜 sarlavhasida sana yo'q bo'lsa, sana None qolardi va
    #         oldingi fayldan qolgan sana ishlatilardi (masalan 30.03).
    # Yechim: Barcha qatorlarni ko'rib birinchi sanani oldindan topamiz.
    # ═══════════════════════════════════════════════════════════════
    global_date = _fayldan_sana_topish(rows)
    if global_date:
        print(f"   📅 Fayl sanasi (oldindan aniqlandi): {global_date.strftime('%d.%m.%Y')}")
    # ═══════════════════════════════════════════════════════════════

    containers = []
    current    = None
    nomalum    = []

    def _flush(c):
        if c and c.get('qatorlar'):
            df = pd.DataFrame(
                c['qatorlar'],
                columns=['Spec_Asl', 'Inventar_Nomi', 'Dona_Soni', 'Topildi']
            )
            df = df[df['Dona_Soni'] > 0].copy()
            c['df'] = df
            del c['qatorlar']

            if not c.get('konteyner_raqam'):
                sana_s = c['sana'].strftime('%d.%m.%Y') if c.get('sana') else '?'
                noraqam_nom = f"NORAQAM_{sana_s.replace('.','_')}_{c['gui_index']}"
                c['konteyner_raqam'] = noraqam_nom
                print(f"   ⚠️  RAQAMSIZ KONTEYNER ({c['gui_index']}) | "
                      f"{sana_s} | {len(df)} tovar")
                print(f"      → {noraqam_nom} sifatida saqlanadi")
                print(f"      备注 da 柜号:XXXX bo'lishi kerak!")

            containers.append(c)

    for ri, row in enumerate(rows):
        row_text = ' '.join(str(c) for c in row if pd.notna(c))

        # Yangi konteyner sarlavhasi: 柜N yoki 小柜N yoki 大柜N
        cell0 = str(row[0]).strip() if pd.notna(row[0]) else ''
        cell1 = str(row[1]).strip() if len(row) > 1 and pd.notna(row[1]) else ''
        gui_cell = ''
        for c in [cell0, cell1]:
            if re.match(r'^[大小]?柜\d*$', c):
                gui_cell = c
                break

        if gui_cell:
            _flush(current)
            sana = None
            # Sanani qatorning BARCHA katakchasidan qidirish
            for cell in row:
                if pd.notna(cell):
                    if hasattr(cell, 'year'):
                        try:
                            sana = datetime(cell.year, cell.month, cell.day)
                            global_date = sana
                            break
                        except Exception:
                            pass
                    cell_str = re.sub(r'\s+', '', str(cell))
                    m = re.search(r'(\d{2})[./](\d{2})[./](\d{4})', cell_str)
                    if m:
                        try:
                            sana = datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
                            global_date = sana
                            break
                        except ValueError:
                            pass
            current = {
                'gui_index':       gui_cell,
                'konteyner_raqam': None,
                'sana':            sana or global_date,  # ← endi global_date to'g'ri
                'tur':             'truba_profil',
                'qatorlar':        [],
                'besshovny':       False,
            }
            continue

        if current is None:
            continue

        # Sana (agar sarlavhada yo'q bo'lsa — qatorlardan qidirish)
        if current['sana'] is None:
            m = re.search(r'(\d{2})[./](\d{2})[./](\d{4})', row_text)
            if m:
                try:
                    current['sana'] = datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
                    global_date = current['sana']
                except ValueError:
                    pass

        # Konteyner raqami
        m_raqam = re.search(r'柜号[：:]\s*([A-Z]{4}\d{7,})', row_text)
        if m_raqam:
            current['konteyner_raqam'] = m_raqam.group(1)

        # Mahsulot qatori
        if spec_col < len(row) and pd.notna(row[spec_col]):
            spec = str(row[spec_col]).strip()

            SKIP_XITOY = [
                '规格', '件数', '合计', '合 计', '板厂重量', '备注', '№',
                'жами', 'jami',
                '配件重量', '外调重量', '外调实心棒',
                '东恒宏钢外调', '家具', '箱子',
                '外调', '装卸费', '打包',
            ]
            # 螺纹 = (Аркон) belgisi — specdan olib, brend_override ga o'tkazamiz
            is_arkon = '螺纹' in spec
            spec = spec.replace('螺纹', '').strip()
            if any(kw.lower() in spec.lower() for kw in SKIP_XITOY) or not spec:
                continue

            if '无缝管' in spec:
                current['besshovny'] = True
                continue

            dona = 0
            if dona_col < len(row) and pd.notna(row[dona_col]):
                try:
                    dona = float(row[dona_col])
                except (ValueError, TypeError):
                    dona = 0

            marka_raw = ''
            if marka_col is not None and marka_col < len(row) and pd.notna(row[marka_col]):
                marka_raw = str(row[marka_col]).strip()

            brend_override = ''
            if is_arkon:
                brend_override = 'Аркон'
            elif current.get('besshovny'):
                if len([p for p in spec.split('*') if p.strip()]) == 3:
                    brend_override = 'Бесшовный'

            nom, topildi = spec_to_nom(
                spec, marka_raw, '', mavjud_nomlar,
                brend_override=brend_override
            )

            if not topildi:
                nomalum.append({
                    'Xitoy_Spec':    spec,
                    'Inventar_Nomi': 'НОМАЪЛУМ',
                    'Konteyner':     current.get('konteyner_raqam') or current['gui_index'],
                    'Sana':          current['sana'].strftime('%d.%m.%Y') if current['sana'] else '?',
                    'Tur':           'truba_profil',
                })

            if dona > 0:
                current['qatorlar'].append((spec, nom, dona, topildi))

    _flush(current)
    return containers, nomalum


# ── List faylini parse qilish ────────────────────────────────────────────────

def _parse_list(df_raw: pd.DataFrame, mavjud_nomlar: set) -> list[dict]:
    """
    List Excel faylidan konteynerlar ro'yxatini qaytaradi.

    Format:
      - Sana sarlavha qatorida (3月20日 yoki 25.03.2026)
      - 材质, 颜色名称, 规格, 数量 ustunlari
      - 备注 ustunida 柜号:XXXX
    """
    rows   = df_raw.values.tolist()
    n_rows = len(rows)

    marka_col  = None
    rang_col   = None
    spec_col   = None
    dona_col   = None
    beixhu_col = None

    for ri in range(min(20, n_rows)):
        for ci, cell in enumerate(rows[ri]):
            if pd.notna(cell):
                s = str(cell).strip()
                if s == '材质'    and marka_col  is None: marka_col  = ci
                if s == '颜色名称' and rang_col   is None: rang_col   = ci
                if s == '规格'    and spec_col   is None: spec_col   = ci
                if s == '数量'    and dona_col   is None: dona_col   = ci
                if s == '备注'    and beixhu_col is None: beixhu_col = ci

    if spec_col is None or dona_col is None:
        raise ValueError(f"List faylida 规格 yoki 数量 topilmadi!")

    containers  = []
    nomalum     = []
    global_date = None

    current_rows  = []
    current_raqam = None
    current_sana  = None

    def _flush_list():
        nonlocal current_rows, current_raqam, current_sana
        if not current_rows:
            current_rows  = []
            current_raqam = None
            return

        df = pd.DataFrame(
            current_rows,
            columns=['Spec_Asl', 'Inventar_Nomi', 'Dona_Soni', 'Topildi']
        )
        df = df[df['Dona_Soni'] > 0].copy()

        if not current_raqam:
            sana_s = current_sana.strftime('%d.%m.%Y') if current_sana else (
                global_date.strftime('%d.%m.%Y') if global_date else '?'
            )
            noraqam_nom = f"NORAQAM_{sana_s.replace('.','_')}_list_{len(containers)+1}"
            print(f"   ⚠️  RAQAMSIZ KONTEYNER (list) | {sana_s} | "
                  f"{len(df)} tovar")
            print(f"      → {noraqam_nom} sifatida saqlanadi")
            print(f"      备注 da 柜号:XXXX bo'lishi kerak!")
            current_raqam = noraqam_nom

        containers.append({
            'gui_index':       current_raqam,
            'konteyner_raqam': current_raqam,
            'sana':            current_sana or global_date,
            'tur':             'list',
            'df':              df,
        })

        current_rows  = []
        current_raqam = None

    for ri, row in enumerate(rows):
        row_text = ' '.join(str(c) for c in row if pd.notna(c))

        # Sanani qidirish — datetime obyekti yoki DD.MM.YYYY string
        for cell in row:
            if pd.notna(cell):
                if hasattr(cell, 'year'):
                    try:
                        s = datetime(cell.year, cell.month, cell.day)
                        global_date  = s
                        current_sana = s
                        break
                    except Exception:
                        pass
                cell_str = re.sub(r'\s+', '', str(cell))
                m = re.search(r'(\d{2})[./](\d{2})[./](\d{4})', cell_str)
                if m:
                    try:
                        s = datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
                        global_date  = s
                        current_sana = s
                        break
                    except ValueError:
                        pass

        # Xitoy oy-kun formati: '3月20日'
        m_oykun = re.search(r'(\d{1,2})月(\d{1,2})日', row_text)
        if m_oykun:
            try:
                yil = global_date.year if global_date else datetime.now().year
                oy  = int(m_oykun.group(1))
                kun = int(m_oykun.group(2))
                s = datetime(yil, oy, kun)
                current_sana = s
                if global_date is None:
                    global_date = s
            except ValueError:
                pass

        # Konteyner raqami
        m_raqam = re.search(r'柜号[：:]\s*([A-Z]{4}\d{7,})', row_text)
        if m_raqam:
            raqam_yangi = m_raqam.group(1)
            if current_raqam is None:
                current_raqam = raqam_yangi
            elif current_raqam != raqam_yangi:
                _flush_list()
                current_raqam = raqam_yangi

        # 合计 qatori → konteyner tugadi
        if '合计' in row_text or '合 计' in row_text:
            if not current_raqam:
                m = re.search(r'柜号[：:]\s*([A-Z]{4}\d{7,})', row_text)
                if m:
                    current_raqam = m.group(1)
            _flush_list()
            continue

        # Mahsulot qatori
        if spec_col < len(row) and pd.notna(row[spec_col]):
            spec = str(row[spec_col]).strip()
            skip = ['规格', '颜色名称', '材质', '数量', '备注', '合计', '克力木', '年']
            if any(kw in spec for kw in skip) or not spec:
                continue
            if not re.search(r'\d.*\*.*\d', spec):
                continue

            dona = 0
            if dona_col < len(row) and pd.notna(row[dona_col]):
                try:
                    dona = float(row[dona_col])
                except (ValueError, TypeError):
                    dona = 0

            marka_raw = ''
            if marka_col is not None and marka_col < len(row) and pd.notna(row[marka_col]):
                marka_raw = str(row[marka_col]).strip()

            rang_raw = ''
            if rang_col is not None and rang_col < len(row) and pd.notna(row[rang_col]):
                rang_raw = str(row[rang_col]).strip()

            nom, topildi = spec_to_nom(spec, marka_raw, rang_raw, mavjud_nomlar)

            if not topildi:
                nomalum.append({
                    'Xitoy_Spec':    spec,
                    'Inventar_Nomi': 'НОМАЪЛУМ',
                    'Konteyner':     current_raqam or '?',
                    'Sana':          current_sana.strftime('%d.%m.%Y') if current_sana else '?',
                    'Tur':           'list',
                })

            current_rows.append((spec, nom, dona, topildi))

    _flush_list()
    return containers, nomalum


# ── Konteynerlarni birlashtirish (raqam + ±2 kun) ───────────────────────────

def _birlashtir(barcha: list[dict]) -> list[dict]:
    """
    Birlashtirish algoritmi:
      1. Raqam bir xil → birlashtir (asosiy qoida)
      2. Raqam bir xil + sana ±2 kun → kichik sana olinadi
      3. Raqam bir xil + sana >2 kun farq → ⚠️ log, baribir birlashtiradi
      4. Raqam yo'q → alohida
    """
    if not barcha:
        return []

    groups  = {}
    noraqam = []

    for c in barcha:
        raqam = c.get('konteyner_raqam')
        if not raqam or raqam.startswith('UNKNOWN'):
            noraqam.append(c)
            continue
        if raqam not in groups:
            groups[raqam] = []
        groups[raqam].append(c)

    result = []

    for raqam, group in groups.items():
        if len(group) == 1:
            result.append(group[0])
            continue

        sanalar = [g['sana'] for g in group if g.get('sana')]
        if len(sanalar) >= 2:
            sanalar_sorted = sorted(sanalar)
            farq = (sanalar_sorted[-1] - sanalar_sorted[0]).days
            if farq > 2:
                print(f"   ⚠️  {raqam}: sana farqi {farq} kun "
                      f"({' | '.join(s.strftime('%d.%m.%Y') for s in sanalar_sorted)}) "
                      f"— baribir birlashtirildi")

        # Sana tanlash qoidasi:
        # truba_profil + list birlashsa → LIST sanasini ol (aniqroq)
        # chunki truba faylida sana ba'zan oldingi konteynerdan o'tib qoladi
        # faqat truba bo'lsa → o'zining sanasini ol
        list_sana   = next((g['sana'] for g in group if g.get('tur') == 'list' and g.get('sana')), None)
        kichik_sana = list_sana if list_sana else (min(sanalar) if sanalar else None)

        dfs = [g['df'] for g in group if 'df' in g and not g['df'].empty]
        merged_df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame(
            columns=['Spec_Asl', 'Inventar_Nomi', 'Dona_Soni', 'Topildi']
        )

        turlar = ' + '.join(sorted(set(g.get('tur', '?') for g in group)))
        result.append({
            'gui_index':       ' + '.join(g['gui_index'] for g in group),
            'konteyner_raqam': raqam,
            'sana':            kichik_sana,
            'tur':             turlar,
            'df':              merged_df,
        })

    result.extend(noraqam)
    return result


# ── Asosiy funksiya ──────────────────────────────────────────────────────────

def parse_china_fayllar(
    fayl_yollari: list[str],
    output_folder: str | None = None,
    mavjud_nomlar: set = None,
    fast: bool = False,
    nomalum_fayl: str = 'NOMALUM.xlsx',
) -> list[dict]:
    """
    Bir yoki bir nechta Xitoy Excel fayllarini parse qiladi.
    Truba/profil va list fayllarini avtomatik aniqlaydi.
    Raqam + ±2 kun bir xil bo'lsa birlashtiradi.

    Parametrlar:
      fayl_yollari  — Excel fayllari ro'yxati
      output_folder — natija papkasi (None bo'lsa saqlamaydi)
      mavjud_nomlar — qoldiq faylidagi mavjud nomlar (1.35 qoidasi uchun)
      fast          — True bo'lsa F_ prefiksi
      nomalum_fayl  — noma'lum speclar fayli

    Qaytaradi: birlashtirgan konteynerlar ro'yxati
    """
    if mavjud_nomlar is None:
        mavjud_nomlar = set()

    barcha_containers = []
    barcha_nomalum    = []
    muvaffaqiyatli    = []

    for fayl in fayl_yollari:
        print(f"\n📂 O'qilmoqda: {os.path.basename(fayl)}")
        try:
            df_raw = pd.read_excel(fayl, header=None)
            tur    = _fayl_turi(df_raw)

            if tur == 'list':
                xatolar = _tekshir_list_fayl(df_raw, fayl)
            else:
                xatolar = _tekshir_truba_fayl(df_raw, fayl)

            if xatolar:
                print(f"   ❌ FAYL FORMATI XATO — o'qilmaydi!")
                for x in xatolar:
                    print(f"      • {x}")
                print(f"   📌 Faylni tekshiring: {os.path.basename(fayl)}")
                continue

            print(f"   Tur: {tur}")

            if tur == 'list':
                containers, nomalum = _parse_list(df_raw, mavjud_nomlar)
            else:
                containers, nomalum = _parse_truba_profil(df_raw, mavjud_nomlar)

            for c in containers:
                print(f"   ✅ {c.get('gui_index','?'):8s} | "
                      f"{c.get('konteyner_raqam','?'):15s} | "
                      f"{c['sana'].strftime('%d.%m.%Y') if c.get('sana') else '?':12s} | "
                      f"{len(c.get('df', [])):3d} tovar")

            barcha_containers.extend(containers)
            barcha_nomalum.extend(nomalum)
            muvaffaqiyatli.append(fayl)

        except Exception as e:
            print(f"   ❌ Xatolik: {e}")
            print(f"   ⚠️  Bu fayl arxivga ko'chirilmaydi — qayta tekshiring!")
            import traceback; traceback.print_exc()

    print(f"\n🔗 Birlashtirish (±2 kun qoidasi)...")
    birlashgan = _birlashtir(barcha_containers)

    truba_map = {}
    list_map  = {}

    for c in barcha_containers:
        raqam = c.get('konteyner_raqam', '?')
        sana  = c['sana'].strftime('%d.%m.%Y') if c.get('sana') else '?'
        if c.get('tur') == 'list':
            list_map[raqam] = sana
        else:
            truba_map[raqam] = sana

    barcha_raqamlar = sorted(set(list(truba_map.keys()) + list(list_map.keys())))

    print(f"\n{'─'*75}")
    print(f"  {'Konteyner':<18} {'Truba/Profil':<15} {'List':<15} Natija")
    print(f"{'─'*75}")

    for raqam in barcha_raqamlar:
        t_sana = truba_map.get(raqam, '—')
        l_sana = list_map.get(raqam,  '—')
        birlashgan_c = next(
            (c for c in birlashgan if c.get('konteyner_raqam') == raqam), None
        )
        tovar_soni = len(birlashgan_c['df']) if birlashgan_c else 0

        if t_sana != '—' and l_sana != '—':
            natija = f"✅ birlashdi  ({tovar_soni} tovar)"
        elif t_sana == '—':
            natija = f"⚠️  faqat list  ({tovar_soni} tovar)"
        else:
            natija = f"⚠️  faqat truba/profil  ({tovar_soni} tovar)"

        print(f"  {raqam:<18} {t_sana:<15} {l_sana:<15} {natija}")

    print(f"{'─'*75}")
    print(f"  Jami: {len(barcha_containers)} ta → {len(birlashgan)} ta konteyner")

    if barcha_nomalum:
        print(f"\n{'─'*60}")
        print(f"⚠️  НОМАЪЛУМ СПЕЦИФИКАЦИЯЛАР ({len(barcha_nomalum)} та):")
        print(f"{'─'*60}")
        for i, n in enumerate(barcha_nomalum, 1):
            print(f"  {i:2d}. {n['Xitoy_Spec']:<25s} "
                  f"({n['Konteyner']}, {n['Sana']}, {n['Tur']})")
        print(f"{'─'*60}")
        print(f"  → {nomalum_fayl} ga saqlandi — to'ldiring va qayta ishga tushiring!")

        nm_df  = pd.DataFrame(barcha_nomalum)
        nm_path = Path(output_folder) / nomalum_fayl if output_folder else Path(nomalum_fayl)
        nm_df.to_excel(str(nm_path), index=False)
    else:
        print(f"\n✅ Barcha speclar aniqlandi — noma'lum yo'q!")

    if output_folder:
        out_dir = Path(output_folder)
        out_dir.mkdir(parents=True, exist_ok=True)

        for c in birlashgan:
            if 'df' not in c or c['df'].empty:
                continue
            sana_str  = c['sana'].strftime('%d.%m.%Y') if c.get('sana') else '01.01.2000'
            raqam     = c.get('konteyner_raqam', 'NORAQAM')
            fast_pref = 'F_' if fast else ''
            fayl_nomi = f"{fast_pref}{raqam}_{sana_str}.xlsx"
            fayl_yoli = out_dir / fayl_nomi

            out_df = c['df'].copy()
            out_df = out_df[out_df['Dona_Soni'] > 0].reset_index(drop=True)
            out_df.insert(0, '№', range(1, len(out_df) + 1))
            out_df = out_df.rename(columns={
                'Inventar_Nomi': 'Товар',
                'Dona_Soni':     'Миқдор',
                'Spec_Asl':      'Xitoy_Spec',
                'Topildi':       'Status',
            })[['№', 'Товар', 'Миқдор', 'Xitoy_Spec', 'Status']]

            out_df.to_excel(str(fayl_yoli), index=False)
            print(f"  💾 {fayl_nomi}  ({len(out_df)} qator)")

    return birlashgan, muvaffaqiyatli


# ── CLI / AUTO-RUN ───────────────────────────────────────────────────────────

if __name__ == '__main__':
    import sys
    import glob

    # ================================================================
    # AUTO-RUN sozlamalari (VS Code Run tugmasi uchun)
    XITOY_PAPKA   = 'Xitoy_Excel'                # Xitoy Excel fayllari shu yerda
    CHIQISH_PAPKA = 'konteynerlar/xitoy_parsed'  # Natija shu yerga chiqadi
    FAST_REJIM    = False                        # True → F_ prefiksi (tezkor konteyner)
    TEST_REJIM    = False                  # True → faqat terminal, hech narsa saqlanmaydi
                                                 # False → haqiqiy saqlash + arxiv
    # ================================================================

    args = sys.argv[1:]

    if args:
        is_fast = '--fast' in args
        args    = [a for a in args if a != '--fast']

        if args and not args[-1].endswith('.xlsx') and not args[-1].endswith('.xls'):
            output_dir = args[-1]
            fayl_args  = args[:-1]
        else:
            output_dir = None
            fayl_args  = args

        fayllar = []
        for a in fayl_args:
            expanded = glob.glob(a)
            fayllar.extend(expanded) if expanded else fayllar.append(a)

    else:
        xitoy_path = Path(SCRIPT_DIR) / XITOY_PAPKA if 'SCRIPT_DIR' in dir() else Path(XITOY_PAPKA)
        fayllar = (
            glob.glob(str(xitoy_path / '*.xlsx')) +
            glob.glob(str(xitoy_path / '*.xls'))
        )
        output_dir = CHIQISH_PAPKA
        is_fast    = FAST_REJIM

        if not fayllar:
            print(f"❌ '{XITOY_PAPKA}' papkasida Excel fayl topilmadi!")
            print(f"   Xitoy fayllarini shu papkaga tashlang va qayta ishga tushiring.")
            sys.exit(1)

    print(f"\n{'='*60}")
    print(f"XITOY KONTEYNER PARSER")
    print(f"{'='*60}")
    print(f"Fayllar:  {len(fayllar)} ta")
    for f in fayllar:
        print(f"  📄 {os.path.basename(f)}")
    print(f"Chiqish:  {output_dir or '(saqlanmaydi)'}")
    print(f"Fast:     {is_fast}")

    if TEST_REJIM and not sys.argv[1:]:
        print(f"\n{'='*60}")
        print(f"🧪 TEST REJIMI — hech narsa saqlanmaydi, arxiv yo'q")
        print(f"{'='*60}")

    birlashgan, muvaffaqiyatli = parse_china_fayllar(
        fayl_yollari=fayllar,
        output_folder=None if (TEST_REJIM and not sys.argv[1:]) else output_dir,
        fast=is_fast,
    )

    if not sys.argv[1:] and not TEST_REJIM:
        import shutil
        from datetime import datetime as _dt

        arxiv_path = Path(XITOY_PAPKA) / 'arxiv'
        arxiv_path.mkdir(exist_ok=True)

        if muvaffaqiyatli:
            print(f"\n📦 Arxivga ko'chirilmoqda → {arxiv_path}/")
            for fayl in muvaffaqiyatli:
                fayl_path = Path(fayl)
                manzil = arxiv_path / fayl_path.name
                if manzil.exists():
                    vaqt  = _dt.now().strftime('%H%M%S')
                    yangi = fayl_path.stem + f'_{vaqt}' + fayl_path.suffix
                    manzil = arxiv_path / yangi
                shutil.move(str(fayl_path), str(manzil))
                print(f"  ✅ {fayl_path.name} → arxiv/{manzil.name}")

        xato_fayllar = [f for f in fayllar if f not in muvaffaqiyatli]
        if xato_fayllar:
            print(f"\n⚠️  Quyidagi fayllar xato bo'lgani uchun Xitoy_Excel/ da qoldi:")
            for f in xato_fayllar:
                print(f"  ❌ {os.path.basename(f)}")

        print(f"\n{'='*60}")
        print(f"✅ TAYYOR! {len(muvaffaqiyatli)} ta fayl arxivga ko'chirildi.")
        if xato_fayllar:
            print(f"⚠️  {len(xato_fayllar)} ta fayl xato — qayta tekshiring!")
        print(f"{'='*60}")

    elif not sys.argv[1:] and TEST_REJIM:
        print(f"\n{'='*60}")
        print(f"🧪 TEST TUGADI!")
        print(f"   Natija yoqsa → TEST_REJIM = False qilib Run bosing")
        print(f"   Muammo bo'lsa → tuzating va qayta Run bosing")
        print(f"{'='*60}")