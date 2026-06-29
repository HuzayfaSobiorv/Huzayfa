"""
vazn_hisobla.py
================
Tovar nomidan (Truba/Profil/List) bitta dona/varaq vaznini (kg) hisoblaydi.
XITOY_DELTA = 0.05 mm — stenka kamaytirish.
"""

import re

ZICHLIK = {
    "201": 7.75,
    "304": 7.93,
    "316": 8.00,
    "321": 7.90,
    "430": 7.70,
    "":    7.85,
}

XITOY_DELTA = 0.05

# 1.35 va 1.45 istisnolar — Xitoy bu qiymatlarni shunday yozadi:
#   Inventar 1,35 -> Xitoy 1.35  (Xitoy 1.30 yo'q)
#   Inventar 1,45 -> Xitoy 1.45  (Xitoy 1.40 yo'q)
#   Inventar 1,4  -> Xitoy 1.35  (normal -0.05)
#   Inventar 1,5  -> Xitoy 1.45  (normal -0.05)
STENKA_EXCEPTION = frozenset({1.35, 1.45})


def _to_float(s: str) -> float:
    cleaned = str(s).replace(',', '.').strip('.')
    if not cleaned:
        raise ValueError(f"Bo'sh qiymat: {s!r}")
    return float(cleaned)


def _stenka_str(val: float) -> str:
    s = f"{val:.2f}".rstrip('0')
    if s.endswith('.'):
        s += '0'
    return s.replace('.', ',')


def xitoy_nomi(name: str) -> str:
    """
    Inventar nomini Xitoy buyurtma nomiga aylantiradi (stenka -XITOY_DELTA).
    Truba/Profil: ст 0,9 -> ст 0,85 | ст 1,4 -> ст 1,35
    List: Лист-0,8 -> Лист-0,75
    ISTISNOLAR: ст 1,35 -> ст 1,35 | ст 1,45 -> ст 1,45
    """
    s = str(name)

    def _sub_stenka(m):
        old_val = _to_float(m.group(1))
        if old_val in STENKA_EXCEPTION:
            return m.group(0)
        new_s = _stenka_str(old_val - XITOY_DELTA)
        return m.group(0).replace(m.group(1), new_s, 1)

    def _sub_list(m):
        old_val = _to_float(m.group(1))
        if old_val in STENKA_EXCEPTION:   # 1.35, 1.45 Лист uchun ham istisnо
            return m.group(0)
        new_s   = _stenka_str(old_val - XITOY_DELTA)
        return m.group(0).replace(m.group(1), new_s, 1)

    # Truba/Profil: "ст N,NN"
    s = re.sub(r'ст\s+([\d,\.]+)', _sub_stenka, s)
    # List: "Лист-N,NN"
    s = re.sub(r'Лист-\s*([\d,\.]+)', _sub_list, s)
    return s


def get_marka(name: str) -> str:
    m = re.search(r'\((\d{3})\s*марка\)', name)
    if m:
        return m.group(1)
    m2 = re.search(r'\b(201|304|316|321|430)\b', name)
    if m2:
        return m2.group(1)
    return ""


def _zichlik(marka: str) -> float:
    return ZICHLIK.get(marka, ZICHLIK[""])


def tovar_vazni(name: str, xitoy: bool = False) -> float | None:
    """
    Tovar nomidan bitta dona (Truba/Profil) yoki bitta varaq (List)
    vaznini kg da qaytaradi. Aniqlanmasa — None.
    xitoy=True: stenka XITOY_DELTA ga kamaytirilgan holda hisoblanadi.
    """
    marka   = get_marka(str(name).strip())
    zichlik = _zichlik(marka)
    s = xitoy_nomi(str(name).strip()) if xitoy else str(name).strip()

    # TRUBA (Бесшовный orasida bo'lsa ham, ( 6 м) space bilan ham ishlaydi)
    m = re.search(
        r'Ф-(\d+(?:[.,]\d+)?)\s*ст\s*([\d,\.]+)[^(]*\(\s*([\d,\.]+)\s*м\)',
        s)
    if m:
        D = _to_float(m.group(1))
        t = _to_float(m.group(2))
        L = _to_float(m.group(3))
        if t <= 0 or t >= D / 2:
            return None
        area_mm2 = 3.14159265 * t * (D - t)
        vol_mm3  = area_mm2 * (L * 1000)
        return round(vol_mm3 * zichlik / 1_000_000, 3)

    # PROFIL (Пр va Пр. ikkalasi ham, ( 6 м) space bilan ham)
    m = re.search(
        r'Пр\.?\s*(\d+(?:[.,]\d+)?)[хx](\d+(?:[.,]\d+)?)\s*ст\s*([\d,\.]+)[^(]*\(\s*([\d,\.]+)\s*м\)',
        s)
    if m:
        a = _to_float(m.group(1))
        b = _to_float(m.group(2))
        t = _to_float(m.group(3))
        L = _to_float(m.group(4))
        perim_mm = 2 * (a + b - 2 * t)
        if perim_mm <= 0:
            return None
        area_mm2 = perim_mm * t
        vol_mm3  = area_mm2 * (L * 1000)
        return round(vol_mm3 * zichlik / 1_000_000, 3)

    # LIST
    m = re.search(r'Лист-\s*([\d,\.]+)\s*\(?(\d+)[хx](\d+)\)?', s)
    if m:
        t   = _to_float(m.group(1))
        en  = float(m.group(2))
        boy = float(m.group(3))
        vol_mm3 = en * boy * t
        return round(vol_mm3 * zichlik / 1_000_000, 3)

    return None
