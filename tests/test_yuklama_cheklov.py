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

import Yuklama_optimal as YO
from Yuklama_optimal import (
    DONA_CHEKLOV_KONTEYNER,
    LIMIT_TRUBA_PROFIL,
    MIN_QATOR_KG_KAT,
    _butun_zaxira_maydami,
    _dona_cheklovi,
    _mayda_qatormi,
    optimallashtir,
)
from vazn_hisobla import tovar_vazni

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


# ── Mayda "bo'lak" (fragment) qoidasi ─────────────────────────────────
# 2026-08-08 (Huzayfa): "ostatkada tayyor mahsulot 2000 ta bor, dastur 1 ta
# konteynerga VAZN TO'G'RILASH uchun 30 ta yuklab qo'ymoqda — bu xato ish".
# ESKI qoida (dona < 20 VA kg < 150) 30-34 donalik ~100 kg bo'lakni
# to'smasdi, chunki dona 20 dan katta edi.
PROF = "Пр. 20х20 ст 0,9 (5,8 м) (201 марка)"
TOLDIRUVCHI = "Ф-32 ст 0,9 (5,8 м) (201 марка)"


class TestMaydaBolak:

    @staticmethod
    def _tolgan_konteyner_ssenariysi():
        """Труба/Профиль slotini ~100 kg qoladigan qilib to'ldiradi —
        ya'ni Пр. 20х20 dan faqat ~34 dona "sig'adigan" holat."""
        v_told = tovar_vazni(TOLDIRUVCHI, xitoy=True)
        n_told = int((LIMIT_TRUBA_PROFIL - 100) / v_told)
        kerak = pd.DataFrame([
            {"Товар": TOLDIRUVCHI, "Холат": "🔴 КРИТИК", "Кам": 99999, "urg_kun": 0},
            {"Товар": PROF,        "Холат": "🔴 КРИТИК", "Кам": 1,     "urg_kun": 0},
        ])
        mavjud = pd.DataFrame([{"Товар": TOLDIRUVCHI, "Миқдор": n_told},
                               {"Товар": PROF,        "Миқдор": 2000}])
        yuklar, qolgan, _, _, _ = optimallashtir(
            kerak, mavjud, abc_map={}, max_yuklar=20)
        return yuklar

    def test_vazn_togrilash_uchun_mayda_bolak_yaratilmaydi(self):
        yuklar = self._tolgan_konteyner_ssenariysi()
        qatorlar = [(i, it) for i, y in enumerate(yuklar, 1)
                    for it in y["items"] if "20х20" in it["tovar"]]
        assert qatorlar, "Пр. 20х20 umuman yuklanmadi"
        for i, it in qatorlar:
            assert it["vazn_kg"] >= MIN_QATOR_KG_KAT["Профиль"], (
                f"konteyner #{i} da kulgili bo'lak: "
                f"{it['dona']} dona / {it['vazn_kg']:.0f} kg"
            )

    def test_hamma_dona_saqlanadi(self):
        yuklar = self._tolgan_konteyner_ssenariysi()
        jami = sum(it["dona"] for y in yuklar for it in y["items"]
                   if "20х20" in it["tovar"])
        assert jami == 2000, f"dona yo'qoldi: {jami}"

    def test_eski_qoida_bugni_bergan_bolardi(self):
        """Eski qoidani vaqtincha tiklab, bug QAYTA HOSIL bo'lishini
        ko'rsatamiz — ya'ni test haqiqatan shu bugni ushlaydi."""
        eski = YO._mayda_qatormi
        try:
            YO._mayda_qatormi = lambda d, v, cat=None: d < 20 and d * v < 150.0
            yuklar = self._tolgan_konteyner_ssenariysi()
            mayda = [it for y in yuklar for it in y["items"]
                     if "20х20" in it["tovar"] and it["vazn_kg"] < 400]
            assert mayda, "eski qoida bilan mayda bo'lak chiqishi kerak edi"
        finally:
            YO._mayda_qatormi = eski


class TestChegaralar:

    @pytest.mark.parametrize("cat,kg_chegara", [
        ("Труба", 400.0), ("Профиль", 400.0), ("Лист", 300.0),
    ])
    def test_kategoriya_chegarasi(self, cat, kg_chegara):
        assert MIN_QATOR_KG_KAT[cat] == kg_chegara
        assert _mayda_qatormi(10, (kg_chegara - 1) / 10, cat) is True
        assert _mayda_qatormi(10, (kg_chegara + 1) / 10, cat) is False

    def test_dona_soni_ozi_ahamiyatsiz(self):
        """ESKI BUG: 30 dona > 20 bo'lgani uchun tekshiruv o'tkazib
        yuborilardi. Endi faqat og'irlik muhim."""
        assert _mayda_qatormi(30, 3.0, "Профиль") is True      # 90 kg
        assert _mayda_qatormi(5, 100.0, "Профиль") is False    # 500 kg

    def test_butun_zaxira_yumshoq_qoida(self):
        """Xitoyda bor-yo'g'i shuncha bo'lsa (sun'iy bo'lak emas) —
        eski, yumshoqroq chegara qoladi (Huzayfa: "agar ostatkada
        shuncha bo'lmasa")."""
        assert _butun_zaxira_maydami(30, 3.0) is False   # 90 kg, lekin hammasi shu
        assert _butun_zaxira_maydami(2, 3.0) is True     # 6 kg — chinakam kulgili
