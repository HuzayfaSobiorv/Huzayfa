"""
real_kunlik_sotuv.py — HAR TOVAR uchun HAQIQIY kunlik sotuvni
tarix/*.xlsx kunlik ombor-qoldig'i fayllaridan hisoblaydigan modul.

2026-07-21 (Huzayfa bilan kelishildi):
=========================================
FON: eski formula (kamomat_engine.zanjir_sim ichida) kunlik sotuvni
"min_zaxira / 30" deb hisoblab kelardi -- bu REAL sotuv emas, min_zaxira'dan
orqaga hisoblangan taxmin edi (o'zi-o'zidan chiqarilgan, "circular"). Real
tarix (tarix/*.xlsx) bilan solishtirilganda bu formula ko'p tovarlarda
haqiqiy sotuvdan 1,6-2,9 baravar (ba'zan undan ham ko'p) katta chiqishi
aniqlandi -- bu esa boshliqning "bot juda ko'p buyurtma beryapti"
shikoyatining asosiy sababi.

QAMROV (Huzayfa qat'iy belgilagan chegaralar):
  1) FAQAT Труба, Профиль, Лист kategoriyalarida qo'llanadi. Boshqa
     kategoriyalar (Баласина, Стойка, Соқка, Чашка va h.k.) eski
     formula (Асосий_Захира/30)da QOLADI -- ular real tarix bilan
     tekshirilmagan, xato xavfi bor (masalan Ф-51/Пр.25х25'da avval
     birinchi urinishda METR/DONA farqini hisobga olmay 5,8 baravar
     xato chiqargan edik -- shu sabab faqat sinovdan o'tgan uchta
     kategoriyaga cheklandi).
  2) Tsex (Цех_Захира) BUTUNLAY ALOHIDA -- fiksirlangan oylik ehtiyoj,
     bu modulga UMUMAN aloqasi yo'q va o'zgarmaydi. Tsexning o'z
     buyurtmasi eski formula (Цех_Захира/30) bilan hisoblanishda davom
     etadi.
  3) MUHIM: agar bitta tovar HAM Асосий, HAM Цех kanaliga tegishli
     bo'lsa (ikkalasida ham min_zaxira > 0), tarix fayllaridagi umumiy
     qoldiq kamayishi IKKALA kanal iste'molini BIRGA ko'rsatadi (tarix
     jismoniy ombordagi UMUMIY qoldiqni yozadi, kanal bo'yicha
     ajratilmagan). Shuning uchun Asosiy uchun real kunlik hisoblanganda
     Tsexning taxminiy ulushi (Цех_Захира/30, ya'ni Tsexning O'ZINING
     eski-uslub kunlik ehtiyoji) tarixdan chiqarilgan umumiy kunlikdan
     AYIRILADI:

         real_kunlik_asosiy = max(0, real_kunlik_jami - sex_kunlik_taxmin)

     Aks holda Tsexning iste'moli "Asosiy juda tez sotilyapti" deb
     noto'g'ri o'qilib, Asosiy buyurtmasini shishirib yuborardi.
  4) Tarixda yetarli ma'lumot bo'lmasa (kam nuqta / qisqa davr / yangi
     tovar -- masalan "Лист-0,55 Кора" hech qachon tarixda uchramagan)
     -- bu tovar natija lug'atiga UMUMAN kiritilmaydi. Chaqiruvchi kod
     (Generate_Asosiy_order.py) topilmasa eski formulaga (Асосий_Захира/30)
     o'zi fallback qiladi -- hech qachon "hisoblay olmadim" deb
     buyurtmani 0 qilib qo'ymaydi.

ISHLATISH:
    from real_kunlik_sotuv import asosiy_kunlik_lugat
    lug'at = asosiy_kunlik_lugat()
    # lug'at: {normallashtirilgan_nom: {"kunlik": 17.45, ...}}
"""

import glob
import logging
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd

logger = logging.getLogger(__name__)

BASE_DIR        = Path(__file__).resolve().parent
TARIX_DIR       = BASE_DIR / "tarix"
MIN_ZAXIRA_FILE = BASE_DIR / "Minimal_zaxiralar" / "Min_Zaxira.xlsx"

# Real hisobga ishonish uchun minimal talablar -- kam bo'lsa fallback.
MIN_KUN_ORALIQ = 14   # kamida shuncha "faol" (tanqissiz) kun bo'lsin
MIN_NUQTA      = 5    # kamida shuncha kunlik yozuv (snapshot) bo'lsin

# 2026-07-21 (Huzayfa bilan kelishildi -- "tanqislik davri" muammosi):
# Tovar tugab/juda kamayib qolgan davrda sotuv TO'XTAYDI -- bu real
# talabni EMAS, ombordagi mol yo'qligini ko'rsatadi (real hodisa:
# "Пр. 40х40 ст 0,9" 27.06dan 21.07gacha ~1 oy DEYARLI O'ZGARMAGAN,
# chunki mol tugagan, talab yo'qolgani uchun emas -- oddiy o'rtacha
# olinsa bu "kam sotiladi" deb NOTO'G'RI xulosaga olib kelardi).
# Yechim: har tovar UCHUN O'ZINING kuzatilgan MAKSIMAL qoldig'idan
# foiz sifatida "tanqislik chegarasi" belgilanadi (universal/qattiq son
# emas -- har tovar o'z shkalasida baholanadi). Interval boshlanishi
# (oldingi kun qiymati) shu chegaradan PAST bo'lsa, o'sha interval
# "faol davr" hisoblanmaydi -- na kamayish, na kun soni hisobga olinadi.
STOCKOUT_FOIZ = 0.15  # tovarning o'z maksimalining 15% dan pastda -- tanqislik

# Faqat shu 3 kategoriyada qo'llanadi (Huzayfa qoidasi, 2026-07-21).
RUXSAT_KATEGORIYA = {"Труба", "Профиль", "Лист"}


# ============================================================
# 1. TARIX FAYLLARINI O'QISH
# ============================================================
def _tarix_fayllar(tarix_dir=None):
    d = Path(tarix_dir) if tarix_dir else TARIX_DIR
    try:
        return sorted(
            glob.glob(str(d / "*.xlsx")),
            key=lambda f: datetime.strptime(Path(f).stem, "%d.%m.%Y"),
        )
    except ValueError:
        # nomi sana formatida bo'lmagan fayl bo'lsa -- shunchaki nom bo'yicha
        return sorted(glob.glob(str(d / "*.xlsx")))


def _upk_songa(upk_raw) -> float:
    """'5,8' -> 5.8 ; None/bo'sh/'1' -> 1.0"""
    if upk_raw is None:
        return 1.0
    try:
        v = float(str(upk_raw).replace(",", "."))
        return v if v > 0 else 1.0
    except (TypeError, ValueError):
        return 1.0


def _barcha_qoldiqlarni_oqi(tarix_dir=None) -> dict:
    """
    Har bir tarix faylini BIR MARTA ochib (125 fayl -- har biri bir marta,
    tovar boshiga qayta ochilmaydi), barcha tovarlarning shu kundagi
    (жami, Упк.мик, Улчови) qiymatini o'qiydi.

    Qaytaradi: {sana: {tovar_nomi: (jami, upk_raw, ulchov)}}
    """
    files = _tarix_fayllar(tarix_dir)
    natija = {}
    for f in files:
        try:
            d = datetime.strptime(Path(f).stem, "%d.%m.%Y")
        except ValueError:
            continue
        try:
            wb = openpyxl.load_workbook(f, data_only=True, read_only=True)
            ws = wb.active
            kun_data = {}
            for row in ws.iter_rows(min_row=6, values_only=True):
                nom = row[0]
                if not nom:
                    continue
                jami = row[8] if len(row) > 8 else None
                if jami is None:
                    continue
                try:
                    jami = float(jami)
                except (TypeError, ValueError):
                    continue
                upk    = row[4] if len(row) > 4 else None
                ulchov = row[6] if len(row) > 6 else None
                kun_data[str(nom).strip()] = (jami, upk, ulchov)
            wb.close()
            natija[d] = kun_data
        except Exception as e:
            logger.warning(f"real_kunlik_sotuv: {f} o'qilmadi: {e}")
            continue
    return natija


# ============================================================
# 2. REAL KUNLIK SOTUV (UMUMIY -- Asosiy+Tsex birga)
# ============================================================
def real_kunlik_jami_hisobla(tarix_dir=None) -> dict:
    """
    HAR bir tovar uchun umumiy (Asosiy+Tsex birgalikda) real kunlik
    sotuvni hisoblaydi. Metr bilan o'lchanadigan tovarlarda
    (Улчови == 'Метр') Упк.мик ga bo'lib donaga o'tkaziladi.

    Metodika (2026-07-21, "tanqislik davri" tuzatilgan versiya):
      1) Kunlik snapshotlar orasidagi FAQAT kamayishlar yig'iladi
         (oshishlar -- konteyner kelishi/tuzatish -- e'tiborga
         olinmaydi; bu usul konteyner qo'shilishlariga tabiiy
         ravishda chidamli).
      2) Har bir interval FAQAT interval boshidagi qiymat shu
         tovarning o'z kuzatilgan maksimalining STOCKOUT_FOIZ (15%)
         dan YUQORI bo'lsagina "faol" (ishonchli) deb hisoblanadi --
         aks holda bu interval "tanqislik davri" (mol tugagan/juda
         kam, sotuv to'xtagan) deb chiqarib tashlanadi -- na kamayish,
         na kun soni umumiy hisobga qo'shilmaydi.
      3) real_kunlik = faol_kamayish / faol_kun (faqat "sog'lom"
         davrlar zichligi asosida -- tanqislik davri buzmaydi).

    Qaytaradi: {tovar_nomi: dict(real_kunlik=.., nuqta=.., kun_oraliq=..)}
    (kun_oraliq bu yerda -- FAOL/hisoblangan kunlar yig'indisi, butun
    kalendar oraliq emas.)
    """
    kunlik_snapshots = _barcha_qoldiqlarni_oqi(tarix_dir)
    sanalar = sorted(kunlik_snapshots.keys())
    if len(sanalar) < 2:
        return {}

    # tovar -> [(sana, jami_dona), ...]
    seriya: dict = {}
    for d in sanalar:
        for nom, (jami, upk, ulchov) in kunlik_snapshots[d].items():
            upk_v = _upk_songa(upk)
            metr_mi = bool(ulchov) and "етр" in str(ulchov)  # 'Метр'/'метр'
            dona = (jami / upk_v) if metr_mi else jami
            seriya.setdefault(nom, []).append((d, dona))

    natija = {}
    for nom, pts in seriya.items():
        if len(pts) < MIN_NUQTA:
            continue
        pts.sort(key=lambda x: x[0])

        max_v = max(v for _, v in pts)
        if max_v <= 0:
            continue
        chegara = max_v * STOCKOUT_FOIZ

        faol_kamayish = 0.0
        faol_kun      = 0
        for i in range(1, len(pts)):
            d_prev, v_prev = pts[i - 1]
            d_curr, v_curr = pts[i]
            gap = (d_curr - d_prev).days
            if gap <= 0:
                continue
            if v_prev <= chegara:
                # tanqislik davri -- interval boshi allaqachon "tugagan"
                # holatda, bu yerdagi (yo'q)lik real talabni ko'rsatmaydi
                continue
            diff = v_curr - v_prev
            if diff < 0:
                faol_kamayish += -diff
            faol_kun += gap

        if faol_kun < MIN_KUN_ORALIQ or faol_kamayish <= 0:
            continue

        real_kunlik = faol_kamayish / faol_kun
        natija[nom] = dict(
            real_kunlik=round(real_kunlik, 3),
            nuqta=len(pts),
            kun_oraliq=faol_kun,
        )
    return natija


# ============================================================
# 3. ASOSIY KANAL UCHUN TAYYOR LUG'AT (Tsex ulushi ayirilgan)
# ============================================================
def asosiy_kunlik_lugat(tarix_dir=None, min_zaxira_file=None) -> dict:
    """
    Asosiy kanal (Generate_Asosiy_order.py, kanal="asosiy") uchun
    ISHLATISHGA TAYYOR real-kunlik lug'atini quradi:
      - faqat Труба/Профиль/Лист kategoriyasidagi tovarlar (Huzayfa
        qoidasi -- boshqa kategoriyalar tekshirilmagan)
      - Tsex ulushi ayirilgan (agar tovar ikkala kanalga ham tegishli
        bo'lsa -- Цех_Захира > 0)
      - tarix yetarli bo'lmagan tovarlar lug'atga UMUMAN kirmaydi
        (chaqiruvchi eski formulaga o'zi fallback qiladi)

    Kalit — common.normalize_product_name() bilan normallashtirilgan
    nom (Generate_Asosiy_order.load_data() dagi bilan bir xil kalit
    orqali moslashtirish uchun).

    Qaytaradi: {normallashtirilgan_nom: dict(kunlik=.., ...)}
    """
    from Generate_Asosiy_order import get_category
    from common import normalize_product_name

    mz_file = Path(min_zaxira_file) if min_zaxira_file else MIN_ZAXIRA_FILE
    sex_map: dict = {}
    if mz_file.exists():
        try:
            mz = pd.read_excel(mz_file, sheet_name="Min_Zaxira")
            mz.columns = [str(c).strip() for c in mz.columns]
            if "Товар" in mz.columns and "Цех_Захира" in mz.columns:
                for _, r in mz.iterrows():
                    nom = r.get("Товар")
                    if pd.isna(nom):
                        continue
                    sex_v = pd.to_numeric(r.get("Цех_Захира"), errors="coerce")
                    if pd.isna(sex_v) or sex_v <= 0:
                        continue
                    sex_map[normalize_product_name(str(nom).strip())] = float(sex_v)
        except Exception as e:
            logger.warning(f"real_kunlik_sotuv: Min_Zaxira.xlsx o'qilmadi: {e}")

    jami = real_kunlik_jami_hisobla(tarix_dir)

    natija = {}
    for nom, info in jami.items():
        if get_category(nom) not in RUXSAT_KATEGORIYA:
            continue
        kalit         = normalize_product_name(nom)
        real_kunlik   = info["real_kunlik"]
        sex_kunlik    = sex_map.get(kalit, 0.0) / 30.0   # Tsexning eski-uslub kunligi
        asosiy_kunlik = max(0.0, real_kunlik - sex_kunlik)
        if asosiy_kunlik <= 0:
            continue
        natija[kalit] = dict(
            kunlik=round(asosiy_kunlik, 2),
            real_kunlik_jami=real_kunlik,
            sex_kunlik_ayirildi=round(sex_kunlik, 2),
            nuqta=info["nuqta"],
            kun_oraliq=info["kun_oraliq"],
        )
    return natija


if __name__ == "__main__":
    lugat = asosiy_kunlik_lugat()
    print(f"Jami {len(lugat)} ta tovar uchun real kunlik sotuv topildi "
          f"(Труба/Профиль/Лист, Tsex ulushi ayirilgan).")
    for nom, info in list(lugat.items())[:10]:
        print(f"  {nom[:55]:55s} kunlik={info['kunlik']:>8.2f}  "
              f"(jami={info['real_kunlik_jami']:.2f}, "
              f"sex_ayirildi={info['sex_kunlik_ayirildi']:.2f}, "
              f"nuqta={info['nuqta']}, kun={info['kun_oraliq']})")
