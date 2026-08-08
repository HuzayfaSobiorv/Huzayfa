# -*- coding: utf-8 -*-
"""
Excel'dagi Қолдиқ / Йўлда ustunlari uchun testlar.

2026-08-08 (Huzayfa, ekran rasmi bilan): "0,8 mat 201 da yo'lda bor, ammo bu
yerda qoldiqni ham yo'ldani ham nol ko'rsatmoqda. Deyarli barcha listlarni
shunday ko'rsatmoqda."

Sabab: `qoldiq_yolda_map` kalitlari faqat normalize_product_name() bilan
qurilardi — u "Лист-" dan keyin bo'shliq qo'yadi ("Лист- 0,8"), inventarda
esa ko'p qatorlar bo'shliqsiz ("Лист-0,8"). Rejadagi nomlar inventarning
aniq shaklida bo'lgani uchun qidiruv topmasdi.

Ishga tushirish (loyiha papkasida):  python -m pytest tests/ -v
"""
import os
import sys

import pandas as pd
import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from yuklatish_rejasi import qoldiq_yolda_map_qur

# Инвентарда bo'shliqSIZ yozilgan (real hayotdagi holat)
INV_NOM = "Лист-0,8 (1220х2440) (Матовый) (201 марка)"
INV = {INV_NOM, "Ф-51 ст 0,9 (5,8 м) (201 марка)"}


def _pb(nom, qoldiq=246, yolda=1135):
    return pd.DataFrame([{"tovar": nom, "qoldiq": qoldiq, "yolda": yolda}])


class TestQoldiqYolda:

    def test_boshliqsiz_inventar_nomi_topiladi(self):
        """ASOSIY BUG: Power BI dagi nom normalize qilinganda bo'shliq
        qo'shiladi, reja esa bo'shliqsiz nom bilan qidiradi."""
        m = qoldiq_yolda_map_qur(_pb(INV_NOM), INV)
        assert m.get(INV_NOM) == (246.0, 1135.0), \
            f"bo'shliqsiz nom topilmadi. Kalitlar: {list(m)}"

    def test_boshliqli_shakl_ham_ishlaydi(self):
        """Eski (bo'shliqli) shakl bilan qidirilsa ham topilishi kerak —
        orqaga muvofiqlik."""
        m = qoldiq_yolda_map_qur(_pb(INV_NOM), INV)
        assert m.get("Лист- 0,8 (1220х2440) (Матовый) (201 марка)") == (246.0, 1135.0)

    def test_truba_buzilmadi(self):
        nom = "Ф-51 ст 0,9 (5,8 м) (201 марка)"
        m = qoldiq_yolda_map_qur(_pb(nom, 1537, 8003), INV)
        assert m.get(nom) == (1537.0, 8003.0)

    def test_inventarda_yoq_tovar_yoqolmaydi(self):
        """Инвентарда bo'lmagan (YANGI) tovar ham map'da qolishi kerak."""
        nom = "Лист-9,9 (1220х2440) (Матовый) (201 марка)"
        m = qoldiq_yolda_map_qur(_pb(nom, 5, 7), INV)
        assert (5.0, 7.0) in m.values(), list(m.items())

    def test_bosh_dataframe(self):
        assert qoldiq_yolda_map_qur(pd.DataFrame(), INV) == {}
        assert qoldiq_yolda_map_qur(None, INV) == {}

    @pytest.mark.parametrize("qoldiq,yolda", [(0, 0), (0, 1135), (246, 0)])
    def test_nol_qiymatlar_saqlanadi(self, qoldiq, yolda):
        """Haqiqiy 0 ni "topilmadi" bilan chalkashtirmaslik kerak."""
        m = qoldiq_yolda_map_qur(_pb(INV_NOM, qoldiq, yolda), INV)
        assert m.get(INV_NOM) == (float(qoldiq), float(yolda))


class TestEskiUsulBugniBerardi:

    def test_eski_usul_topolmasdi(self):
        """Tuzatish haqiqatan kerak bo'lganini ko'rsatadi: faqat _norm()
        bilan qurilgan map bo'shliqsiz nomni TOPMAYDI."""
        from common import normalize_product_name as _norm
        eski = {_norm(INV_NOM): (246.0, 1135.0)}
        assert INV_NOM not in eski, "bug qayta hosil bo'lmadi"
        assert _norm(INV_NOM) != INV_NOM
