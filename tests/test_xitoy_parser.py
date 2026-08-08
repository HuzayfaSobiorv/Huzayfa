# -*- coding: utf-8 -*-
"""
Xitoy ostatka/yuklatish parseri uchun regression testlar.

2026-08-08 sessiyasida topilgan REAL buglar (Huzayfa: "konteyner yuklatish
ishlamayapti, ko'plab mahsulotlarni tanimayapti"). Har bir test bug'ni
QAYTA HOSIL qiladigan aniq kirish bilan yozilgan — ya'ni tuzatish orqaga
qaytarilsa, test yiqiladi.

Ishga tushirish (loyiha papkasida):  python -m pytest tests/ -v
"""
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from parsers import (
    _china_spec_to_inventar,
    _inventar_snap,
    _kanonik_nom,
    _parse_list_xitoy,
)


# ── BUG A: stenka raqami ichidagi bo'shliq ─────────────────────────────
# "Metalmart_Jadvallar_Jamlanma.xlsx" (2026-08-08) faylida stenka qo'lda
# terilgani uchun bo'shliq bilan yozilgan: "Φ 16 cT 0, 65".
# ESKI regex `([\d,\.]+)` bo'shliqda to'xtab, faqat "0," ni olardi →
# _yaxlitla_stenka("0,") = "0,0".
class TestStenkaBoshliq:

    @pytest.mark.parametrize("spec,kutilgan", [
        # (bo'shliqli variant,                 kutilgan stenka)
        ("Φ 16 cT 0, 65", "0,65"),
        ("Φ 25 cT 1. 20", "1,2"),
        ("Φ 32 cT 0. 85", "0,9"),
        ("Φ 38 cT 1. 20", "1,2"),
        ("Φ 38 cT 1. 35", "1,35"),
        ("Φ 51 cT 0. 85", "0,9"),
    ])
    def test_truba_boshliqli_stenka(self, spec, kutilgan):
        nom = _china_spec_to_inventar(spec, "5.8M")
        assert nom is not None, f"{spec!r} umuman tanilmadi"
        assert f"ст {kutilgan} " in nom, f"{spec!r} → {nom!r}"

    @pytest.mark.parametrize("spec,kutilgan", [
        ("KB 20x20 CT 0. 65", "0,65"),
        ("KB 40x40 CT 1. 95", "2,0"),
        ("KB 50x25 CT 1. 05", "1,1"),
    ])
    def test_profil_boshliqli_stenka(self, spec, kutilgan):
        nom = _china_spec_to_inventar(spec, "5.8M")
        assert nom is not None, f"{spec!r} umuman tanilmadi"
        assert f"ст {kutilgan} " in nom, f"{spec!r} → {nom!r}"

    def test_boshliqsiz_variant_buzilmadi(self):
        """Eski (bo'shliqsiz) format ilgarigidek ishlashi shart."""
        assert _china_spec_to_inventar("φ51 cT 0.85", "5.8M") == \
            "Ф-51 ст 0,9 (5,8 м) (201 марка)"
        assert _china_spec_to_inventar("KB 20x20 CT 0.65", "5.8M") == \
            "Пр. 20х20 ст 0,65 (5,8 м) (201 марка)"

    def test_ikki_xil_stenka_qoshilib_ketmasin(self):
        """Bug'ning eng xavfli oqibati: "Φ 38 cT 1. 20" va "Φ 38 cT 1. 35"
        ikkalasi ham "ст 1,0" bo'lib, IKKI XIL mahsulot bitta qatorga
        qo'shilib ketardi (456 + 297 = 753)."""
        a = _china_spec_to_inventar("Φ 38 cT 1. 20", "5.8M")
        b = _china_spec_to_inventar("Φ 38 cT 1. 35", "5.8M")
        assert a != b, f"ikki xil stenka bitta nomga tushdi: {a!r}"

    def test_notogri_mahsulotga_ulanmasin(self):
        """Eng jim xavf: "Φ 25 cT 1. 20" → "Ф-25 ст 1,0" bo'lardi, bu nom
        inventarda MAVJUD — ya'ni bot xatoni ko'rsatmasdan BOSHQA
        mahsulotni yuklatish rejasiga qo'yardi."""
        nom = _china_spec_to_inventar("Φ 25 cT 1. 20", "5.8M")
        assert "ст 1,0 " not in nom, f"noto'g'ri mahsulotga ulandi: {nom!r}"


# ── BUG B: Лист markasi 材质 ustunidan o'qilmasdi ──────────────────────
# "Xiaoshou_Qingdan_2026-08-08.xlsx" (销售清单) formatida 品号 ustuni YO'Q,
# material 材质 ustunida ("201直板"/"304直板"). Marka topilmagach nom
# "(201 марка)" qo'shimchasisiz yasalardi va 201/304 QO'SHILIB ketardi.
def _sotuv_royxati_rows():
    """销售清单 formatidagi minimal jadval (品号 ustunisiz)."""
    return [
        ("佛山... 销售清单", "", "", "", "", "", "", ""),
        ("客户: ...", "", "", "", "", "单号: CH2608080108", "", ""),
        ("订单号", "材质", "表面颜色", "规格", "负差客标", "单位", "数量", "备注"),
        ("SD1", "201直板", "砂板", "0.80*1219*2438", "-6c", "张", 1700, ""),
        ("SD1", "304直板", "砂板", "0.80*1219*2438", "-6c", "张", 100, ""),
        ("SD1", "201直板", "8K钛金", "1.00*1219*2438", "-6c", "张", 100, ""),
    ]


class TestListMarka:

    def test_marka_cai_zhi_dan_oqiladi(self):
        known, _ = _parse_list_xitoy(_sotuv_royxati_rows())
        assert any("(201 марка)" in n for n in known), \
            f"201 марка topilmadi: {list(known)}"
        assert any("(304 марка)" in n for n in known), \
            f"304 марка topilmadi: {list(known)}"

    def test_201_va_304_qoshilib_ketmasin(self):
        """Bir xil qalinlik/o'lcham/rang, lekin har xil marka — ALOHIDA
        qator bo'lishi shart (ilgari 1700 + 100 = 1800 bo'lib ketardi)."""
        known, _ = _parse_list_xitoy(_sotuv_royxati_rows())
        m201 = [v for n, v in known.items() if "0,8" in n and "(201 марка)" in n]
        m304 = [v for n, v in known.items() if "0,8" in n and "(304 марка)" in n]
        assert m201 == [1700], f"201 miqdori noto'g'ri: {m201}"
        assert m304 == [100], f"304 miqdori noto'g'ri: {m304}"

    def test_rang_va_olcham(self):
        # Nom oxirida inventarning ASL yozuvi bilan almashtirilishi mumkin
        # ("Лист- 1,0" ↔ "Лист-1,0"), shuning uchun kanonik solishtiramiz.
        known, _ = _parse_list_xitoy(_sotuv_royxati_rows())
        kutilgan = _kanonik_nom("Лист-1,0 (1220х2440) (Голд) (201 марка)")
        assert any(_kanonik_nom(n) == kutilgan for n in known), list(known)

    def test_inventarning_asl_yozuvi_qaytariladi(self):
        """BUG D (2026-08-08, oxirigacha simulyatsiyada topilgan): nom
        to'g'ri yasalsa ham inventar bilan mos kelmasdi, chunki
        normalize_product_name() "Лист-" dan keyin DOIM bo'shliq qo'yadi
        ("Лист- 0,8"), inventarda esa ko'p qatorlar bo'shliqSIZ
        ("Лист-0,8"). Parser endi inventardagi ASL yozuvni qaytaradi."""
        import parsers
        eski_set, eski_kmap = parsers._inventar_set_cache, parsers._inventar_kanonik_cache
        try:
            parsers._inventar_set_cache = {
                "Лист-0,8 (1220х2440) (Матовый) (201 марка)",   # bo'shliqSIZ
            }
            parsers._inventar_kanonik_cache = None
            known, _ = _parse_list_xitoy(_sotuv_royxati_rows())
            assert "Лист-0,8 (1220х2440) (Матовый) (201 марка)" in known, list(known)
            # Inventarda yo'q tovar o'zgarishsiz (bo'shliqli) qolishi kerak
            assert any(n.startswith("Лист- ") for n in known), list(known)
        finally:
            parsers._inventar_set_cache = eski_set
            parsers._inventar_kanonik_cache = eski_kmap

    def test_pin_hao_li_format_buzilmadi(self):
        """品号 ustuni BOR bo'lgan eski format ilgarigidek ishlashi shart."""
        rows = [
            ("品号", "颜色", "规格", "数量"),
            ("304ABC", "砂板", "1.20*1219*2438", 50),
        ]
        known, _ = _parse_list_xitoy(rows)
        assert any("(304 марка)" in n for n in known), list(known)


# ── BUG C: qalinlik snap'i UZUNLIKni ham almashtirardi ────────────────
# _inventar_snap 0,65/1,35/1,45 stenkalarda nomni birinchi "(" gacha kesib
# qidirardi — qavsdan keyingi qism esa UZUNLIK. Natijada Xitoy 5,8 м
# tovari inventardagi 6 м tovarga jimgina ulanardi.
class TestSnapUzunlikniOzgartirmasin:

    INV = {
        "Ф-38 ст 1,4 (6 м) (201 марка)",
        "Ф-38 ст 1,35 (6 м) (304 марка)",
        "Ф-16 ст 0,7 (5,8 м) (201 марка)",
        "Пр. 40х40 ст 1,4 (5,8 м) (201 марка)",
    }

    def test_uzunlik_kesib_otilmasin(self):
        """5,8 м tovar 6 м tovarga ULANMASLIGI shart (Huzayfa qoidasi:
        5,8 topilmasa — bu YANGI mahsulot)."""
        nom = "Ф-38 ст 1,35 (5,8 м) (201 марка)"
        assert _inventar_snap(nom, self.INV) == nom

    def test_stenka_snap_ishlashda_davom_etsin(self):
        """Uzunlik AYNI bo'lsa, qalinlik yaxlitlash ilgarigidek ishlaydi."""
        assert _inventar_snap("Ф-16 ст 0,65 (5,8 м) (201 марка)", self.INV) == \
            "Ф-16 ст 0,7 (5,8 м) (201 марка)"
        assert _inventar_snap("Пр. 40х40 ст 1,35 (5,8 м) (201 марка)", self.INV) == \
            "Пр. 40х40 ст 1,4 (5,8 м) (201 марка)"

    def test_topilmasa_asl_nom(self):
        nom = "Ф-99 ст 1,35 (5,8 м) (201 марка)"
        assert _inventar_snap(nom, self.INV) == nom
