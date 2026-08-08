# -*- coding: utf-8 -*-
"""
Konteyner ichidagi DONA cheklovi uchun testlar (Yuklama_optimal.py).

2026-08-08 (Huzayfa): "51 dan maksimalniy 800 ta yuklansin, sababi kubi
sig'may qoladi". Optimizator faqat og'irlikni hisoblaydi, hajmni emas —
shu sabab Ф-51 bitta konteynerga 1827 donagacha tushib ketardi.

Ishga tushirish (loyiha papkasida):  python -m pytest tests/ -v
"""
import os
import sys

import pandas as pd
import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from Yuklama_optimal import (
    DONA_CHEKLOV_KONTEYNER,
    _dona_cheklovi,
    optimallashtir,
)

F51 = "Ф-51 ст 0,9 (5,8 м) (201 марка)"
LIMIT = 800


def _optimallashtir(tovarlar: dict):
    """tovarlar: {nom: dona} → (yuklar, qolgan)"""
    kerak = pd.DataFrame(
        [{"Товар": n, "Холат": "🔴 КРИТИК", "Кам": d, "urg_kun": 0}
         for n, d in tovarlar.items()]
    )
    mavjud = pd.DataFrame(
        [{"Товар": n, "Миқдор": d} for n, d in tovarlar.items()]
    )
    yuklar, qolgan, _, _, _ = optimallashtir(kerak, mavjud, abc_map={}, max_yuklar=20)
    return yuklar, qolgan


def _dona(yuk, nom_bolagi):
    return sum(it["dona"] for it in yuk["items"] if nom_bolagi in it["tovar"])


# ── Shablon moslashuvi ────────────────────────────────────────────────
class TestShablon:

    @pytest.mark.parametrize("nom", [
        "Ф-51 ст 0,9 (5,8 м) (201 марка)",
        "Ф-51 ст 0,9 (6 м) (201 марка)",      # uzunlik farqi yo'q
        "Ф-51 ст 0,9 (6 м) (304 марка)",      # marka farqi yo'q
        "Ф-51 ст 1,4 (6 м) (201 марка)",      # qalinlik farqi yo'q
        "Ф-51 ст 3,0 Бесшовный (6 м) (304 марка)",
        "(Аркон) Ф-51 ст 0,9 (6 м) (201 марка)",   # brend prefiksi bilan
    ])
    def test_f51_cheklovga_tushadi(self, nom):
        assert _dona_cheklovi(nom) == (list(DONA_CHEKLOV_KONTEYNER)[0], LIMIT)

    @pytest.mark.parametrize("nom", [
        "Ф-16 ст 0,7 (5,8 м) (201 марка)",
        "Ф-5 ст 0,9 (6 м) (201 марка)",
        "Ф-510 ст 0,9 (6 м) (201 марка)",     # 51 bilan boshlansa ham — boshqa tovar
        "Пр. 51х51 ст 0,9 (6 м) (201 марка)",
        "Лист- 0,8 (1220х2440) (Матовый) (201 марка)",
    ])
    def test_boshqa_tovarlar_cheklanmaydi(self, nom):
        assert _dona_cheklovi(nom) is None


# ── Optimizatordagi haqiqiy xatti-harakat ─────────────────────────────
class TestOptimallashtir:

    def test_bir_konteynerda_800_dan_oshmasin(self):
        yuklar, _ = _optimallashtir({F51: 6472})
        assert yuklar, "hech qanday konteyner yasalmadi"
        for i, y in enumerate(yuklar, 1):
            assert _dona(y, "Ф-51") <= LIMIT, \
                f"konteyner #{i} da {_dona(y, 'Ф-51')} dona — chegara {LIMIT}"

    def test_hamma_tovar_yuklanadi(self):
        """Cheklov tovarni YO'QOTMASLIGI kerak — ortiqchasi keyingi
        konteynerga o'tadi, "yuklanmadi" ro'yxatiga tushmaydi."""
        yuklar, qolgan = _optimallashtir({F51: 6472})
        jami = sum(_dona(y, "Ф-51") for y in yuklar)
        jami += sum(q["dona"] for q in qolgan if "Ф-51" in q["tovar"])
        assert jami == 6472, f"dona yo'qoldi/ko'paydi: {jami}"

    def test_kam_miqdor_bitta_konteynerda(self):
        """Chegaradan kam bo'lsa — bo'linmaydi."""
        yuklar, _ = _optimallashtir({F51: 500})
        toliq = [y for y in yuklar if _dona(y, "Ф-51") > 0]
        assert len(toliq) == 1, f"{len(toliq)} ta konteynerga bo'lindi"
        assert _dona(toliq[0], "Ф-51") == 500

    def test_har_xil_qalinlik_birga_hisoblanadi(self):
        """Ф-51 ning turli qalinliklari BITTA guruh — konteyner bo'yicha
        yig'indisi 800 dan oshmasligi kerak."""
        yuklar, _ = _optimallashtir({
            "Ф-51 ст 0,9 (5,8 м) (201 марка)": 600,
            "Ф-51 ст 1,4 (6 м) (201 марка)": 600,
        })
        for i, y in enumerate(yuklar, 1):
            assert _dona(y, "Ф-51") <= LIMIT, \
                f"konteyner #{i}: {_dona(y, 'Ф-51')} dona"

    def test_cheklovsiz_tovar_buzilmadi(self):
        """Cheklovga tushmaydigan tovar ilgarigidek — faqat og'irlik
        chegarasi bilan yuklanadi (800 dona qoidasi qo'llanmaydi)."""
        nom = "Ф-16 ст 0,7 (5,8 м) (201 марка)"
        yuklar, _ = _optimallashtir({nom: 3000})
        eng_kop = max(_dona(y, "Ф-16") for y in yuklar)
        assert eng_kop > LIMIT, \
            f"cheklovsiz tovarga ham 800 chegarasi qo'llandi ({eng_kop})"
