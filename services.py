"""
services.py — Biznes logika: inventar, kamomat, buyurtma, konteyner
"""
import difflib
import json
import logging
import re
from datetime import datetime
from io import BytesIO
from pathlib import Path

import pandas as pd

logger = logging.getLogger(__name__)

from config import (
    BASE_DIR, DATA_FILE, BOT_HOLAT_DIR, VARAQLAR,
    CAT_SHEET, AKSESSUAR_KATS, get_inv, get_kont,
    KONTEYNER_TARIX_FILE, XITOY_PARSED_DIR,
)
from common import normalize_product_name

# Lokal aliaslar — config cache funksiyalari
_get_inv  = get_inv
_get_kont = get_kont


# ── Og'irlik lookup (vazn_lookup.xlsx — Power BI dan mustaqil) ───────────────
# Fayl: BASE_DIR / "vazn_lookup.xlsx"
# Sheet: "Вазн" | Ustunlar: Товар номи (A) | 1 дона вазни кг (B)
# Qoida: stenka -0.05 (Xitoy), ISTISNO: 1.35 va 1.45 stenka ayrilmaydi
_vazn_cache: dict | None = None
VAZN_FILE = BASE_DIR / "vazn_lookup.xlsx"

def vazn_map_yuklash(force: bool = False) -> dict:
    """
    vazn_lookup.xlsx fayldan tovar_nomi → og'irlik (kg) lug'atini qaytaradi.
    Power BI yangilanganda ham bu fayl o'zgarmaydi — mustaqil saqlanadi.
    """
    global _vazn_cache
    if _vazn_cache is not None and not force:
        return _vazn_cache
    try:
        import openpyxl
        if not VAZN_FILE.exists():
            logger.warning(f"vazn_lookup.xlsx topilmadi: {VAZN_FILE}")
            _vazn_cache = {}
            return _vazn_cache
        wb = openpyxl.load_workbook(VAZN_FILE, data_only=True, read_only=True)
        ws = wb.active
        result = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[0]:
                continue
            tovar = str(row[0]).strip()
            vazn  = row[1]
            if isinstance(vazn, (int, float)) and vazn > 0:
                result[tovar] = float(vazn)
        wb.close()
        _vazn_cache = result
        logger.info(f"vazn_map_yuklash: {len(result)} ta tovar og'irligi yuklandi.")
        return result
    except Exception as e:
        logger.error(f"vazn_map_yuklash xato: {e}")
        _vazn_cache = {}
        return {}


def tovar_vazni_pb(tovar_nomi: str) -> float | None:
    """
    vazn_lookup.xlsx dan tovar og'irligini qaytaradi. Topilmasa — None.
    """
    vmap = vazn_map_yuklash()
    return vmap.get(str(tovar_nomi).strip())


# ── Whitelist ────────────────────────────────────────────────────────────────
_WHITELIST_FILE = BOT_HOLAT_DIR / "whitelist.json"


def whitelist_yuklash() -> set[int]:
    """Ruxsat berilgan foydalanuvchilar ID larini qaytaradi."""
    if not _WHITELIST_FILE.exists():
        return set()
    try:
        return set(json.loads(_WHITELIST_FILE.read_text(encoding="utf-8")))
    except Exception:
        return set()


def whitelist_saqlash(ids: set[int]) -> None:
    _WHITELIST_FILE.write_text(json.dumps(sorted(ids)), encoding="utf-8")


def whitelist_qosh(uid: int) -> bool:
    """ID ni whitelist ga qo'shadi. True — yangi qo'shildi, False — allaqachon bor."""
    wl = whitelist_yuklash()
    if uid in wl:
        return False
    wl.add(uid)
    whitelist_saqlash(wl)
    return True


def whitelist_ochir(uid: int) -> bool:
    """ID ni whitelist dan o'chiradi. True — o'chirildi, False — topilmadi."""
    wl = whitelist_yuklash()
    if uid not in wl:
        return False
    wl.discard(uid)
    whitelist_saqlash(wl)
    return True


# ── Konteyner qo'shish tarixi (fayl o'chsa ham unutilmasin) ───────────────────
# MUHIM: kalit ISO emas, ISO+sana ("ISO|sana" shaklida) — chunki jismoniy
# konteyner raqamlari (ISO) dunyoda qayta-qayta ishlatiladi va bitta ISO
# boshqa sanada butunlay YANGI yuk bilan qaytib kelishi mumkin. Faqat aynan
# bir xil ISO+sana (ya'ni xuddi shu yetkazib berish) qayta bloklanadi.

def konteyner_tarix_kalit(iso: str, sana: str) -> str:
    return f"{iso}|{sana}"


def konteyner_tarix_olish() -> set[str]:
    """Bir marta tasdiqlangan konteynerlar (ISO+sana kalit shaklida)
    to'plamini qaytaradi. xitoy_parsed papkasidagi fayl keyinchalik
    o'chirilgan bo'lsa ham, shu kalit qayta "yangi" deb qo'shilib
    ketmasligi uchun ishlatiladi.

    Birinchi chaqiriqda (tarix fayli hali yo'q bo'lsa) hozirgi
    xitoy_parsed papkasidagi BARCHA konteynerlar bilan bir martalik
    "boshlang'ich to'ldirish" qilinadi — shu orqali oldindan (bu tarix
    tizimi yaratilishidan oldin) qo'shilgan konteynerlar ham himoyalanadi."""
    if not KONTEYNER_TARIX_FILE.exists():
        boshlangich = set()
        if XITOY_PARSED_DIR.exists():
            for f in XITOY_PARSED_DIR.glob("*.xlsx"):
                stem = f.stem[:-2] if f.stem.endswith("_D") else f.stem
                # 2026-07-13: "F_" (tezkor/aksessuar) prefiksi ham olib
                # tashlanadi -- aks holda (masalan "F_AksessuarKont_26.06.2026")
                # noto'g'ri iso="F" bo'lib chiqib, bu konteyner keyinchalik
                # dedup tekshiruvida "ko'rinmas" bo'lib qolardi (konteyner_qosh.py
                # dagi bir xil tuzatish bilan izchil).
                if stem.startswith("F_"):
                    stem = stem[2:]
                iso, _, sana = stem.partition("_")
                if iso:
                    boshlangich.add(konteyner_tarix_kalit(iso, sana))
        try:
            KONTEYNER_TARIX_FILE.write_text(
                json.dumps(sorted(boshlangich), ensure_ascii=False), encoding="utf-8"
            )
        except Exception as e:
            logger.error(f"konteyner_tarix boshlang'ich to'ldirishda xato: {e}")
        return boshlangich
    try:
        return set(json.loads(KONTEYNER_TARIX_FILE.read_text(encoding="utf-8")))
    except Exception:
        return set()


def konteyner_tarix_qoshish(konteynerlar: list) -> None:
    """Tasdiqlangan konteynerlarni doimiy tarixga qo'shadi.
    `konteynerlar` — [{"iso":.., "sana":..}, ...] yoki (iso, sana) juftliklari."""
    if not konteynerlar:
        return
    tarix = konteyner_tarix_olish()
    for k in konteynerlar:
        if isinstance(k, dict):
            iso, sana = k["iso"], k["sana"]
        else:
            iso, sana = k
        tarix.add(konteyner_tarix_kalit(iso, sana))
    try:
        KONTEYNER_TARIX_FILE.write_text(
            json.dumps(sorted(tarix), ensure_ascii=False), encoding="utf-8"
        )
    except Exception as e:
        logger.error(f"konteyner_tarix_qoshish yozishda xato: {e}")


def kirish_ruxsati(uid: int) -> bool:
    """Foydalanuvchi kirish huquqiga ega ekanligini tekshiradi."""
    from config import ADMIN_IDS
    if uid in ADMIN_IDS:
        return True
    return uid in whitelist_yuklash()


def vazn_lookup_yangilash(yangi_vazn: dict) -> int:
    """
    Xitoy faylidan olingan yangi tovar vaznlarini vazn_lookup.xlsx ga qo'shadi.
    FAQAT mavjud bo'lmagan tovarlar qo'shiladi — eski yozuvlar o'ZGARMAYDI.
    Qaytaradi: qo'shilgan tovarlar soni.
    """
    if not yangi_vazn:
        return 0
    global _vazn_cache
    try:
        import openpyxl
        mavjud = vazn_map_yuklash()
        yangilar = {k: v for k, v in yangi_vazn.items()
                    if str(k).strip() and k not in mavjud
                    and isinstance(v, (int, float)) and v > 0}
        if not yangilar:
            return 0
        if VAZN_FILE.exists():
            wb = openpyxl.load_workbook(VAZN_FILE)
        else:
            wb = openpyxl.Workbook()
            wb.active.append(["Товар номи", "1 дона вазни кг"])
        ws = wb.active
        for nom, vazn in sorted(yangilar.items()):
            ws.append([nom, round(float(vazn), 3)])
            logger.info(f"vazn_lookup: yangi tovar qo'shildi — {nom}: {vazn} kg")
        wb.save(VAZN_FILE)
        # Keshni yangilaymiz
        _vazn_cache = None
        vazn_map_yuklash()
        return len(yangilar)
    except Exception as e:
        logger.error(f"vazn_lookup_yangilash xato: {e}")
        return 0


def xlsx_mavjud() -> bool:
    return DATA_FILE.exists()


def _num(df: pd.DataFrame, col: str):
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)


def inventar_olish(kanal: str | None = None) -> pd.DataFrame:
    from common import keraksizmi
    try:
        df = _get_inv().copy()
        for col in ["Қолдиқ", "Мин_Захира", "Фарқ", "Кун_Етади", "Йўлда_Жами"]:
            _num(df, col)
        if kanal and "Тур" in df.columns:
            if kanal == "sex":
                df = df[df["Тур"] == "ЦЕХ🏭"]
            elif kanal in ("asosiy", "osh"):
                df = df[df["Тур"] != "ЦЕХ🏭"]
        if "Товар" in df.columns:
            df = df[~df["Товар"].apply(keraksizmi)].copy()
        return df.reset_index(drop=True)
    except Exception as e:
        logger.error(f"inventar_olish: {e}")
        return pd.DataFrame()


def _kamomat_filter(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "Холат" not in df.columns:
        return pd.DataFrame()
    return df[df["Холат"].isin(["🔴 КРИТИК", "🟡 ПАСТ"])].copy()


def kamomat_olish(kanal: str) -> pd.DataFrame:
    try:
        return _kamomat_filter(inventar_olish(kanal)).reset_index(drop=True)
    except Exception as e:
        logger.error(f"kamomat_olish: {e}")
        return pd.DataFrame()


def kritiklar_olish(kanal: str, limit: int = 15) -> pd.DataFrame:
    df = inventar_olish(kanal)
    if df.empty or "Холат" not in df.columns:
        return pd.DataFrame()
    df = df[df["Холат"] == "🔴 КРИТИК"].copy()
    if df.empty:
        return df
    sort_cols = [c for c in ["Кун_Етади", "Фарқ", "Товар"] if c in df.columns]
    return df.sort_values(sort_cols).head(limit).reset_index(drop=True)


def kritiklar_text(df: pd.DataFrame, lang: str) -> str:
    if df.empty:
        return "Bugun kritik tovar yo'q." if lang == "lat" else "Бугун критик товар йўқ."
    title = "Bugungi eng xavfli kritiklar:" if lang == "lat" else "Бугунги энг хавфли критиклар:"
    lines = [title]
    for i, row in df.iterrows():
        tovar = str(row.get("Товар", "?"))
        qoldiq = int(row.get("Қолдиқ", 0))
        min_z = int(row.get("Мин_Захира", 0))
        kun = row.get("Кун_Етади", "")
        farq = int(row.get("Фарқ", 0))
        lines.append(
            f"{i + 1}. {tovar}\n"
            f"   Qoldiq: {qoldiq} | Min: {min_z} | Farq: {farq} | Kun: {kun}"
        )
    return "\n".join(lines)


# ── Kirill/lotin bir xil deb hisoblovchi normalizatsiya (umumiy qidiruv) ──────
_CYR2LAT = {
    'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'yo',
    'ж': 'j', 'з': 'z', 'и': 'i', 'й': 'y', 'к': 'k', 'л': 'l', 'м': 'm',
    'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u',
    'ф': 'f', 'х': 'x', 'ц': 'ts', 'ч': 'ch', 'ш': 'sh', 'щ': 'sh',
    'ъ': '', 'ы': 'i', 'ь': '', 'э': 'e', 'ю': 'yu', 'я': 'ya',
    'қ': 'q', 'ғ': 'g', 'ў': 'o', 'ҳ': 'h',
}


def _translit(s: str) -> str:
    """Kirillni lotinga o'giradi (harf-baharf); lotin harflar o'zgarishsiz qoladi."""
    return "".join(_CYR2LAT.get(ch, ch) for ch in s.lower())


def _qidiruv_normalize(s: str) -> str:
    """Qidiruv uchun: kirill/lotin, katta/kichik harf va '.'/',' farqini
    yo'qotadi; so'z chegaralarini (bo'sh joy) saqlab qoladi — pastda
    tokenlarga ajratish (fuzzy qidiruv) uchun kerak."""
    s = _translit(str(s))
    s = s.replace(".", ",")
    s = re.sub(r"[\'ʻʼ`’]", "", s)
    s = re.sub(r"[^a-z0-9,]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _fuzzy_score(terms: list, name_tokens: list) -> float:
    """Har bir so'rov so'zi uchun nom ichidagi eng yaqin so'zni topib,
    o'rtacha o'xshashlik darajasini qaytaradi (0..1)."""
    if not terms or not name_tokens:
        return 0.0
    total = 0.0
    for term in terms:
        best = max(
            (difflib.SequenceMatcher(None, term, tok).ratio() for tok in name_tokens),
            default=0.0,
        )
        total += best
    return total / len(terms)


def qidiruv_olish(query: str, kanal: str | None = None, limit: int = 12) -> pd.DataFrame:
    """
    Umumiy (kategoriyasiz) qidiruv — butun inventar bo'yicha.
    Kirill va lotin yozuvi farqi hisobga olinmaydi (masalan "труба" va "truba"
    bir xil natija beradi), '.' va ',' ham bir xil qabul qilinadi.
    Aniq mos kelmasa — xato/yozuv farqiga chidamli (taxminiy/fuzzy) eng
    yaqin natijalarni qaytaradi.
    """
    df = inventar_olish(kanal)
    if df.empty or "Товар" not in df.columns:
        return pd.DataFrame()

    terms = [_qidiruv_normalize(t) for t in re.split(r"\s+", query.strip()) if t.strip()]
    terms = [t for t in terms if t]
    if not terms:
        return pd.DataFrame()

    names_norm = df["Товар"].astype(str).apply(_qidiruv_normalize)
    mask = pd.Series(True, index=df.index)
    for term in terms:
        mask &= names_norm.str.contains(re.escape(term), na=False)

    out = df[mask].copy()

    # ── Aniq mos kelmasa — so'z darajasida taxminiy (fuzzy) qidiruv ─────────
    if out.empty:
        name_tokens = names_norm.str.split(" ")
        scores = name_tokens.apply(lambda toks: _fuzzy_score(terms, toks))
        best = scores[scores >= 0.6].sort_values(ascending=False)
        if not best.empty:
            out = df.loc[best.index[:limit]].copy()

    if out.empty:
        return out
    status_rank = {"🔴 КРИТИК": 1, "🟡 ПАСТ": 2, "🟢 НОРМА": 3, "МЕЁР ЙЎҚ": 4}
    if "Холат" in out.columns:
        out["_rank"] = out["Холат"].map(status_rank).fillna(9)
    else:
        out["_rank"] = 9
    sort_cols = [c for c in ["_rank", "Кун_Етади", "Товар"] if c in out.columns]
    return (
        out.sort_values(sort_cols)
        .head(limit)
        .drop(columns=["_rank"], errors="ignore")
        .reset_index(drop=True)
    )


def qidiruv_text(query: str, df: pd.DataFrame, lang: str) -> str:
    if df.empty:
        return f"Natija topilmadi: {query}" if lang == "lat" else f"Натижа топилмади: {query}"
    title = f"Qidiruv: {query}" if lang == "lat" else f"Қидирув: {query}"
    lines = [title]
    for i, row in df.iterrows():
        tovar = str(row.get("Товар", "?"))
        holat = str(row.get("Холат", ""))
        qoldiq = int(row.get("Қолдиқ", 0))
        yolda = int(row.get("Йўлда_Жами", 0))
        min_z = int(row.get("Мин_Захира", 0))
        kun = row.get("Кун_Етади", "")
        lines.append(
            f"{i + 1}. {tovar}\n"
            f"   {holat} | Qoldiq: {qoldiq} | Yo'lda: {yolda} | Min: {min_z} | Kun: {kun}"
        )
    return "\n".join(lines)


def grafik_qidirish(query: str, kat: str, kanal: str) -> pd.DataFrame:
    """
    "51 -> 0,9 -> 5,8 -> 201" yoki "51 0,9 5,8 201" formatini parse qilib tovar topadi.
    kat: truba | profil | list | bal
    '.' va ',' bir xil, lotin 'x'/'X' va kirill 'х' bir xil qabul qilinadi.
    """
    # Normalize: nuqtani vergulga, lotin x ni kirill х ga
    query = query.replace('.', ',').replace('x', 'х').replace('X', 'х')

    # -> bilan yoki bo'sh joy bilan ham ishlaydi
    if '->' in query:
        parts = [p.strip() for p in query.split('->')]
    else:
        parts = [p.strip() for p in query.split()]
    df = inventar_olish(kanal)
    if df.empty or "Товар" not in df.columns:
        return pd.DataFrame()

    names = df["Товар"].astype(str)
    mask  = pd.Series(True, index=df.index)

    def _q(s):
        return re.escape(s)  # allaqachon vergul formatida

    # 2026-07-11 (tuzatildi): raqamli qidiruvlar oldin ANCHOR'siz substring
    # edi -- "0,6" qidirilsa "0,65"/"0,61" ham, "Ф-19" qidirilsa "Ф-190" ham
    # (bo'lsa) aralashib chiqardi, chunki "0,6" harfma-harf "0,65" ichida
    # ham bor. Endi raqamdan keyin YANA raqam kelmasligi tekshiriladi
    # (?!\d) -- shunda "0,6" faqat "0,6" bilan tugagan joyda mos keladi,
    # "0,65" ga mos kelmaydi.
    def _qn(s):
        return re.escape(s.strip()) + r'(?!\d)'

    if kat == "truba":
        if len(parts) >= 1:
            mask &= names.str.contains(f'Ф-{_qn(parts[0])}', na=False, case=False)
        if len(parts) >= 2:
            mask &= names.str.contains(f'ст {_qn(parts[1])}', na=False, case=False)
        if len(parts) >= 3:
            mask &= names.str.contains(_qn(parts[2]), na=False, case=False)
        if len(parts) >= 4:
            mask &= names.str.contains(parts[3].strip(), na=False, case=False)

    elif kat == "profil":
        # "20 20 0,7 6 201" yoki "20х20 0,7 6 201" — ikkalasi ishlaydi
        # Birinchi ikki qism raqam bo'lsa (masalan "20" "20") → "20х20" deb birlashtir
        idx = 0
        if len(parts) >= 2:
            try:
                float(parts[0].replace(',', '.')); float(parts[1].replace(',', '.'))
                # Ikkalasi raqam → o'lcham ikki alohida qism
                olcham = f"{parts[0]}х{parts[1]}"
                idx = 2
            except ValueError:
                olcham = parts[0].strip()
                idx = 1
        elif len(parts) >= 1:
            olcham = parts[0].strip()
            idx = 1
        else:
            olcham = ""
        if olcham:
            mask &= names.str.contains(_qn(olcham), na=False, case=False)
        if len(parts) > idx:
            mask &= names.str.contains(f'ст {_qn(parts[idx])}', na=False, case=False)
        if len(parts) > idx + 1:
            mask &= names.str.contains(_qn(parts[idx + 1]), na=False, case=False)
        if len(parts) > idx + 2:
            mask &= names.str.contains(parts[idx + 2].strip(), na=False, case=False)

    elif kat == "list":
        if len(parts) >= 1:
            mask &= names.str.contains(f'Лист.*{_qn(parts[0])}', na=False, case=False)
        if len(parts) >= 2:
            mask &= names.str.contains(parts[1].strip(), na=False, case=False)
        if len(parts) >= 3:
            mask &= names.str.contains(parts[2].strip(), na=False, case=False)

    elif kat == "bal":
        q = query.strip()
        if re.match(r'^\d+$', q):
            num = q.zfill(2)
            mask &= names.str.contains(f'Баласина-{num}', na=False, case=False)
        else:
            mask &= names.str.lower().str.contains(re.escape(q.lower()), na=False)

    elif kat == "stoyka":
        mask &= names.str.contains('Стойка', na=False, case=False)
        q = query.strip()
        if q:
            # "01" → "№01", "34" → "№34"
            num = re.sub(r'[^0-9,]', '', q)
            if num:
                mask &= names.str.contains(re.escape(num), na=False, case=False)

    elif kat == "chas":
        mask &= names.str.contains('Чашка', na=False, case=False)
        if len(parts) == 1:
            p0 = parts[0].strip()
            # Boshidagi raqam/o'lcham bo'yicha filtr
            mask &= names.str.contains(f'^{re.escape(p0)}', na=False, regex=True, case=False)
        elif len(parts) >= 2:
            # "40 40" → profil chashkasi "40х40-Чашка"
            olcham = f"{parts[0]}х{parts[1]}"
            mask &= names.str.contains(re.escape(olcham), na=False, case=False)

    elif kat == "kuz":
        mask &= names.str.contains('Кузикорин', na=False, case=False)
        q = query.strip()
        if q:
            mask &= names.str.contains(f'^{re.escape(q)}', na=False, regex=True, case=False)

    elif kat == "shar":
        mask &= names.str.contains('Шар', na=False, case=False)
        mask &= ~names.str.contains('Шаркона|Шарнир|Шаршара', na=False, case=False)
        q = query.strip()
        if q:
            if len(parts) >= 2:
                # "4 4" yoki "4х4" → profil sharlari
                olcham = f"{parts[0]}х{parts[1]}"
                mask &= names.str.contains(f'^{re.escape(olcham)}', na=False, regex=True, case=False)
            else:
                # "51", "76" — diametr bo'yicha
                mask &= names.str.contains(f'^{re.escape(q)}', na=False, regex=True, case=False)

    elif kat == "sokka":
        mask &= names.str.contains('Сокка', na=False, case=False)
        mask &= ~names.str.contains('Баласина', na=False, case=False)
        if len(parts) == 1:
            p0 = parts[0].strip()
            mask &= names.str.contains(f'^{re.escape(p0)}', na=False, regex=True, case=False)
        elif len(parts) >= 2:
            # "90 25" → "Сокка 90х25" yoki "90х25"
            olcham = f"{parts[0]}х{parts[1]}"
            mask &= names.str.contains(re.escape(olcham), na=False, case=False)

    elif kat == "oyna":
        mask &= names.str.contains('Ойна держатель', na=False, case=False)
        q = query.strip()
        num = re.sub(r'[^0-9]', '', q)
        if num:
            mask &= names.str.contains(f'№-{num.zfill(2)}', na=False, case=False)

    return df[mask].copy().reset_index(drop=True)


def buyurtma_yuklash(kanal: str) -> dict | None:
    p = BOT_HOLAT_DIR / f"buyurtma_{kanal}.json"
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except:
        return None


def buyurtma_saqlash(kanal: str, items: list):
    """
    Yangi tasdiqlangan buyurtmani mavjud tasdiqlangan bilan BIRLASHTIRADI.
    Bir tovar uchun miqdorlar yig'iladi — ustiga yozilmaydi.
    Masalan: 1-kun 619 ta + 3-kun 50 ta = jami 669 ta saqlanadi.
    """
    p = BOT_HOLAT_DIR / f"buyurtma_{kanal}.json"

    # Mavjud tasdiqlangan buyurtmalarni yuklash
    mavjud = buyurtma_yuklash(kanal)
    eski_map: dict[str, float] = {}
    if mavjud and mavjud.get("buyurtmalar"):
        for eski in mavjud["buyurtmalar"]:
            tov = str(eski.get("tovar", "")).strip()
            miq = float(eski.get("miqdor", 0))
            if tov:
                eski_map[tov] = eski_map.get(tov, 0) + miq

    # Yangilarini qo'shish (yig'ish)
    for yangi in items:
        tov = str(yangi.get("tovar", "")).strip()
        miq = float(yangi.get("miqdor", 0))
        if tov and miq > 0:
            eski_map[tov] = eski_map.get(tov, 0) + miq

    # Birlashtirilgan ro'yxat
    jami_items = [
        {"tovar": tov, "miqdor": miq}
        for tov, miq in eski_map.items()
        if miq > 0
    ]

    p.write_text(json.dumps(
        {"sana":        datetime.now().strftime("%d.%m.%Y %H:%M"),
         "kanal":       kanal,
         "buyurtmalar": jami_items},
        ensure_ascii=False, indent=2), encoding="utf-8")


def pending_saqlash(kanal: str, user_id: int, items: list):
    """Tasdiqlash kutilayotgan buyurtmani diskka saqlaydi (bot qayta ishga tushsa yo'qolmasin)."""
    p = BOT_HOLAT_DIR / f"pending_{kanal}_{user_id}.json"
    p.write_text(json.dumps(
        {"kanal": kanal, "user_id": user_id, "items": items,
         "sana": datetime.now().strftime("%d.%m.%Y %H:%M")},
        ensure_ascii=False, indent=2), encoding="utf-8")


def pending_yuklash(kanal: str, user_id: int) -> list | None:
    """Diskdan pending zakaz yuklaydi. Topilmasa None."""
    p = BOT_HOLAT_DIR / f"pending_{kanal}_{user_id}.json"
    if not p.exists():
        return None
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        return data.get("items")
    except:
        return None


def pending_tozala(kanal: str, user_id: int):
    """Pending zakaz faylini o'chiradi."""
    p = BOT_HOLAT_DIR / f"pending_{kanal}_{user_id}.json"
    if p.exists():
        p.unlink()


def draft_saqlash(kanal: str, tovarlar: list):
    """Excel yaratilganda tovar → to'liq nom mappingini saqlaydi.
    buyurtma_tekshir shu fayldan lookup qiladi — inventar o'zgarsa ham ishlaydi."""
    p = BOT_HOLAT_DIR / f"draft_{kanal}.json"
    p.write_text(json.dumps(
        {"sana": datetime.now().strftime("%d.%m.%Y %H:%M"),
         "tovarlar": tovarlar},
        ensure_ascii=False, indent=2), encoding="utf-8")


def draft_yuklash(kanal: str) -> list | None:
    p = BOT_HOLAT_DIR / f"draft_{kanal}.json"
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8")).get("tovarlar")
    except:
        return None


def _kun_oldin(sana_str: str) -> str:
    """'10.06.2026 14:23' → 'Bugun', 'Kecha', '3 kun oldin', yoki '15 kun oldin ⚠️'"""
    try:
        d = datetime.strptime(sana_str, "%d.%m.%Y %H:%M")
        delta = (datetime.now() - d).days
        if delta == 0:
            return "Bugun"
        elif delta == 1:
            return "Kecha"
        elif delta <= 13:
            return f"{delta} kun oldin"
        else:
            return f"{delta} kun oldin ⚠️"
    except Exception:
        return sana_str


def xitoy_yuklash(kanal: str) -> dict | None:
    p = BOT_HOLAT_DIR / f"xitoy_{kanal}.json"
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except:
        return None


def xitoy_saqlash(kanal: str, xitoy_map: dict, ombor_map: dict | None = None,
                   vazn_map: dict | None = None):
    """
    xitoy_map  — K ustun (jami zakaz, buyurtma hisoblash uchun)
    ombor_map  — L ustun (tayyor/ready, yuklatish rejasi uchun)
    vazn_map   — {tovar_nomi: 1_dona_kg} Xitoy faylidan olingan haqiqiy vazn
    """
    p = BOT_HOLAT_DIR / f"xitoy_{kanal}.json"
    data = {
        "sana":     datetime.now().strftime("%d.%m.%Y %H:%M"),
        "kanal":    kanal,
        "tovarlar": xitoy_map,
    }
    if ombor_map:
        data["ombor"] = ombor_map
    if vazn_map:
        data["vazn"] = vazn_map
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def buyurtma_tozala(kanal: str):
    """Xitoy ostatka yuklanganda chaqiriladi — tasdiqlangan buyurtma endi
    K ustunda aks etgan, shu sababli alohida hisobga olish kerak emas."""
    p = BOT_HOLAT_DIR / f"buyurtma_{kanal}.json"
    if p.exists():
        p.unlink()


def zakaz_preview_text(kanal: str, items: list, lang: str) -> str:
    ch = t(lang, CH_KEY[kanal])
    by_varaq: dict = {}
    for i in items:
        v = i.get("varaq", "?")
        by_varaq[v] = by_varaq.get(v, 0) + 1
    tafsilot = "".join(f"\n  • {v}: {c} ta" for v, c in by_varaq.items())

    # KRITIK tovarlar soni — inventardan cross-reference
    kritik_satri = ""
    try:
        inv_df = inventar_olish(kanal)
        if not inv_df.empty and "Товар" in inv_df.columns and "Холат" in inv_df.columns:
            kritik_set = set(inv_df[inv_df["Холат"] == "🔴 КРИТИК"]["Товар"].astype(str))
            n_kritik   = sum(1 for i in items if str(i.get("tovar","")).strip() in kritik_set)
            if n_kritik > 0:
                if lang == "cyr":
                    kritik_satri = f"\n🔴 Шундан *{n_kritik} та* КРИТИК ҳолатдаги товар."
                else:
                    kritik_satri = f"\n🔴 Shundan *{n_kritik} ta* KRITIK holatdagi tovar."
    except Exception:
        pass

    # Oldingi tasdiqlangan buyurtma borligini tekshirish — akkumulyatsiya ogohlantirish
    oldingi = buyurtma_yuklash(kanal)
    ogoh = ""
    if oldingi and oldingi.get("buyurtmalar"):
        n_oldin = len(oldingi["buyurtmalar"])
        sana    = oldingi.get("sana", "?")
        if lang == "cyr":
            ogoh = f"\n\n⚠️ *Диққат:* {kanal} каналида аввал *{n_oldin} та* товар тасдиқланган ({sana}).\nЯнгиси устига *қўшилади* — бекор қилиш учун Созламалар → Буюртмани тозалаш."
        else:
            ogoh = f"\n\n⚠️ *Diqqat:* {kanal} kanalida avval *{n_oldin} ta* tovar tasdiqlangan ({sana}).\nYangisi ustiga *qo'shiladi* — bekor qilish uchun Sozlamalar → Buyurtmani tozalash."

    return t(lang, "zakaz_preview").format(
        ch=ch, n=len(items), tafsilot=tafsilot
    ) + kritik_satri + ogoh


def kamomat_stats(kanal: str) -> dict:
    df  = kamomat_olish(kanal)
    if df.empty:
        return {"n": 0, "b": 0, "p": 0}
    buy     = buyurtma_yuklash(kanal)
    ordered = {i["tovar"] for i in buy.get("buyurtmalar",[])} if buy else set()
    tovs    = df["Товар"].tolist() if "Товар" in df.columns else []
    b       = sum(1 for x in tovs if x in ordered)
    return {"n": len(df), "b": b, "p": len(df) - b}



def asosiy_styled_excel_yarat(xitoy_ostatka: dict | None = None,
                               kanal: str = "asosiy") -> BytesIO:
    from Generate_Asosiy_order import load_data, calculate, build

    # Agar xitoy_ostatka berilmagan bo'lsa — JSON fayldan avtomatik o'qiymiz
    mavjud = xitoy_yuklash(kanal)
    if xitoy_ostatka is None:
        if mavjud and mavjud.get("tovarlar"):
            xitoy_ostatka = mavjud["tovarlar"]
    # 2026-07-13 (Huzayfa: "K va L ustunlarini Excelda ham ko'rsatib
    # turish kerak"): Tayyor (L) ustuni ham xuddi shu xitoy_{kanal}.json
    # faylida ("ombor" kaliti) saqlanadi -- ILGARI bu yerda umuman
    # o'qilmasdi, faqat Zakaz (K, "tovarlar") buyurtmani kamaytirish
    # uchun ishlatilardi. Endi ikkalasi ham Excelga (E=Zakaz, F=Tayyor)
    # yoziladi -- admin har bir tovar uchun Xitoydagi ostatkani buyurtma
    # qatori yonida to'g'ridan-to'g'ri ko'radi, alohida fayl ochib
    # qidirishga hojat qolmaydi.
    ombor_map = (mavjud or {}).get("ombor") or {}

    _df, _kont_map = load_data(kanal=kanal)
    df_calc = calculate(_df, _kont_map)

    # 1. Xitoy ostatka K ustuni ayiriladi (Xitoyda buyurtma berilgan + tayyorlanayotgan)
    if xitoy_ostatka and not df_calc.empty:
        def _adjust_xitoy(row):
            ayir = float(xitoy_ostatka.get(str(row["tovar"]).strip(), 0))
            return max(0, int(row["buyurtma"]) - int(ayir))
        df_calc["buyurtma"] = df_calc.apply(_adjust_xitoy, axis=1)
        df_calc = df_calc[df_calc["buyurtma"] > 0].copy()

    # 2. Tasdiqlangan buyurtma ayiriladi (allaqachon buyurilgan, lekin yo'lda emas)
    #    Agar tasdiqlangan hamma yetishmovchilikni qoplasa → buyurtma = 0 → Excel bo'sh
    tasdiq = buyurtma_yuklash(kanal)
    if tasdiq and tasdiq.get("buyurtmalar") and not df_calc.empty:
        tasdiq_map = {}
        for item in tasdiq["buyurtmalar"]:
            tov = str(item.get("tovar", "")).strip()
            miq = float(item.get("miqdor", 0))
            if tov:
                tasdiq_map[tov] = tasdiq_map.get(tov, 0) + miq
        if tasdiq_map:
            oldin = len(df_calc)
            def _adjust_tasdiq(row):
                ayir = tasdiq_map.get(str(row["tovar"]).strip(), 0)
                return max(0, int(row["buyurtma"]) - int(ayir))
            df_calc["buyurtma"] = df_calc.apply(_adjust_tasdiq, axis=1)
            df_calc = df_calc[df_calc["buyurtma"] > 0].copy()
            keyin = len(df_calc)
            logger.info(
                f"[{kanal}] tasdiqlangan ayirish: {oldin} → {keyin} ta "
                f"({oldin-keyin} ta o'chirildi, {len(tasdiq_map)} ta tasdiqlangan)"
            )

    if df_calc.empty:
        return None  # Buyurtma kerak emas — chaqiruvchi xabar beradi

    # E/F ustunlari uchun — xuddi shu tovar nomi bo'yicha Zakaz/Tayyor
    # qiymatlarini qo'shib qo'yamiz (FAQAT ko'rsatish uchun, hisoblashga
    # ta'sir qilmaydi — "buyurtma" ustuni yuqorida allaqachon tuzatilgan).
    df_calc["zakaz"] = df_calc["tovar"].apply(
        lambda t: xitoy_ostatka.get(str(t).strip(), 0) if xitoy_ostatka else 0
    )
    df_calc["tayyor"] = df_calc["tovar"].apply(
        lambda t: ombor_map.get(str(t).strip(), 0) if ombor_map else 0
    )

    # Draft tovarlarni saqlaymiz — buyurtma_tekshir shu ro'yxatdan lookup qiladi.
    # Inventar keyinchalik o'zgarsa ham xato bo'lmaydi.
    draft_saqlash(kanal, df_calc["tovar"].tolist())

    wb  = build(df_calc)
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


def draft_excel_yarat(kanal: str, xitoy_ostatka: dict | None = None):
    """Barcha kanallar uchun bir xil styled Excel. None → buyurtma kerak emas."""
    return asosiy_styled_excel_yarat(xitoy_ostatka, kanal=kanal)


def buyurtma_tekshir(fayl_bytes: bytes, kanal: str = "asosiy"):
    """
    Tasdiqlangan buyurtma Excel'ini tekshiradi.
    Generate_Asosiy_order.build() formati: "Tovar nomi" | "Uzunlik" | "Buyurtma"
    """
    import openpyxl
    from Generate_Asosiy_order import strip_length, get_length

    try:
        wb = openpyxl.load_workbook(BytesIO(fayl_bytes))
    except Exception as e:
        return False, f"Faylni ochishda xato: {type(e).__name__}", None

    # Lookup manbayi: AVVALO draft faylidan (Excel yaratilganda saqlangan).
    # Sabab: inventar o'zgarishi mumkin, lekin Excel dagi tovar nomlari o'sha paytgi
    # drafdtga mos. Draft yo'q bo'lsa inventardan fallback qilinadi.
    #
    # MUHIM: Excel col1 da rich_tovar_name() → strip_length(xitoy_nomi(original)) yoziladi.
    # xitoy_nomi() stenka qiymatini STENKA_DELTA ga kamaytiradi (0.26 → 0.21).
    # Shuning uchun lookup kaliti strip_length(xitoy_nomi(full_name)) bo'lishi kerak,
    # aks holda mos kelinmaydi va tasdiqlangan buyurtma saqlanmaydi.
    try:
        from vazn_hisobla import xitoy_nomi as _xitoy_nomi
    except ImportError:
        def _xitoy_nomi(n): return n   # fallback

    lookup = {}
    draft = draft_yuklash(kanal)
    if draft:
        for full_name in draft:
            full_name = str(full_name).strip()
            # Excel col1 da ko'rsatiladigan nom (stenka kamaytirgan)
            excel_nomi = strip_length(_xitoy_nomi(full_name)).strip()
            uzun       = get_length(full_name).strip()
            lookup[(excel_nomi, uzun)] = full_name
            # fallback: uzunsiz ham topilsin
            if (excel_nomi, "") not in lookup:
                lookup[(excel_nomi, "")] = full_name
    else:
        inv = inventar_olish(kanal)
        if inv.empty or "Товар" not in inv.columns:
            return False, "Инвентар ma'lumoti topilmadi", None
        for tovar in inv["Товар"].dropna().astype(str):
            nomi = strip_length(tovar).strip()
            uzun = get_length(tovar).strip()
            lookup[(nomi, uzun)] = tovar

    items = []
    found_any = False

    for ws in wb.worksheets:
        header_row = None
        for r in range(1, min(ws.max_row, 5) + 1):
            v = ws.cell(row=r, column=1).value
            if v is not None and str(v).strip().lower() == "tovar nomi":
                header_row = r
                break
        if header_row is None:
            continue
        found_any = True

        for r in range(header_row + 1, ws.max_row + 1):
            cell1 = ws.cell(row=r, column=1)
            nomi = cell1.value
            if nomi is None or str(nomi).strip() == "":
                continue
            nomi = str(nomi).strip()
            uzun_val = ws.cell(row=r, column=2).value
            uzun = "" if uzun_val is None else str(uzun_val).strip()
            buy_val = ws.cell(row=r, column=3).value
            if buy_val in (None, ""):
                continue
            # Avvalo cell comment dan asl inventar nomini o'qiymiz.
            # Sabab: ст 1,35 (inventar) va ст 1,4 (inventar) ikkalasi Xitoy
            # Excelda "ст 1,35" ko'rinadi — lookup collision'dan qochish uchun
            # write_product() asl nomni comment sifatida yozadi.
            if cell1.comment and cell1.comment.text and cell1.comment.text.strip():
                tovar = cell1.comment.text.strip()
            else:
                # Eski Excel (commentsiz) uchun fallback
                tovar = lookup.get((nomi, uzun)) or lookup.get((nomi, ""))
            if tovar is None:
                continue
            try:
                m = float(buy_val)
            except (ValueError, TypeError):
                return False, f"'{ws.title}' varaqida raqam bo'lmagan buyurtma qiymati", None
            if m > 0:
                items.append({"tovar": tovar, "miqdor": m, "varaq": ws.title})

    if not found_any:
        return False, (
            "Hech qaysi varaqda 'Tovar' ustuni topilmadi.\n"
            "Faylda quyidagi varaqlar bor: "
            + ", ".join(str(s) for s in wb.sheetnames)
        ), None
    return True, None, items
