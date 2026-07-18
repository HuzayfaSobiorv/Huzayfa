# -*- coding: utf-8 -*-
"""
zanjir_sim (2026-07-14 order-up-to versiyasi) testlari.

Kelishuv (Huzayfa, 2026-07-14):
  - Gorizont = KELISH_KUNI (55) — yangi buyurtma yetib kelguncha davr
  - Trigger: gorizont ichida qoldiq min dan pastga tushsa
  - Hajm: min + BUYURTMA_SIKL_KUN(30) kunlik savdo darajasigacha to'ldirish
"""
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from kamomat_engine import zanjir_sim
from common import KELISH_KUNI, KUNLIK_SOTUV_BOLISH, BUYURTMA_SIKL_KUN

# Standart misol: min=3000 → kunlik = 100
MIN = 3000
KUNLIK = MIN / KUNLIK_SOTUV_BOLISH  # 100


class TestGorizont:
    """1-tuzatish: 30 kun emas, 55 kun oldinga qarash."""

    def test_erta_konteynerdan_keyingi_tanqislik_korinadi(self):
        """ESKI BUG: konteyner 10-kun kelsa, sim 40-kunda to'xtab,
        40-55 kunlardagi tanqislikni ko'rmasdi (taklif 2500 chiqardi).
        YANGI: 55-kungacha qaraladi → taklif 7000."""
        s = zanjir_sim(qoldiq=2500, min_z=MIN, konteynerlar=[(10, 2000)])
        # 2026-07-18 (lost-sales clamp): 55-kun prognoz endi 0 da to'xtaydi
        # (minus "yo'qotilgan sotuv" buyurtmaga qo'shilmaydi).
        # nishon: 3000 + 100*30 = 6000 → taklif = 6000 - 0 = 6000
        assert s["taklif"] == 6000
        assert s["min_nuqta"] == -1000   # min_nuqta clamp'siz (axborot uchun)
        assert s["xavf"] == "KRITIK"

    def test_uzilish_kuni_gorizont_ichida(self):
        s = zanjir_sim(qoldiq=5900, min_z=MIN, konteynerlar=[])
        # (5900-3000)/100 = 29-kun min chizig'i kesiladi
        assert s["uzilish_kun"] == 29

    def test_kech_konteyner_hisobga_olinadi_lekin_simga_kirmaydi(self):
        """55+ kunda keladigan konteyner: uzilishni oldini olmaydi,
        lekin buyurtma hajmidan ayiriladi (ikki marta buyurmaslik)."""
        s = zanjir_sim(qoldiq=3000, min_z=MIN, konteynerlar=[(60, 5000)])
        # 55-kun: 3000 - 5500 → clamp bilan 0 → uzilish bor
        assert s["xavf"] == "KRITIK"
        # 2026-07-18 (lost-sales clamp): nishon 6000 - (0 + 5000) = 1000
        assert s["taklif"] == 1000


class TestOrderUpTo:
    """2-tuzatish: mayda 'titrash' o'rniga oyiga 1 marta yirik buyurtma."""

    def test_yuqori_zaxirada_buyurtma_yoq(self):
        """qoldiq 55 kundan ortiq yetsa — 0 (eski kod 50-200 chiqarardi)."""
        s = zanjir_sim(qoldiq=9000, min_z=MIN, konteynerlar=[])
        assert s["taklif"] == 0
        assert s["xavf"] == "NORMA"

    def test_mayda_taklif_chiqmaydi(self):
        """ESKI BUG: qoldiq=5800 da 200 talik mikro-taklif chiqardi.
        YANGI: trigger otilgach bir yo'la yirik buyurtma."""
        s = zanjir_sim(qoldiq=5800, min_z=MIN, konteynerlar=[])
        # 55-kun: 300 → taklif = 6000 - 300 = 5700 (mayda emas!)
        assert s["taklif"] == 5700
        assert s["taklif"] >= KUNLIK * BUYURTMA_SIKL_KUN  # kamida ~1 oylik

    def test_trigger_chegarasi(self):
        """Trigger aynan: qoldiq < min + 55 kunlik savdo (= 8500)."""
        assert zanjir_sim(8500, MIN, [])["taklif"] == 0   # roppa-rosa yetadi
        assert zanjir_sim(8400, MIN, [])["taklif"] > 0    # 100 ta kam — trigger

    def test_buyurtmadan_keyin_uzoq_jim(self):
        """Buyurtma berilgandan keyin (yo'lda deb hisoblasak) tovar
        yana chiqmasligi kerak — titrash tugadi."""
        s1 = zanjir_sim(qoldiq=5800, min_z=MIN, konteynerlar=[])
        # buyurtma yo'lga chiqdi deb faraz: 55-kunda keladi
        s2 = zanjir_sim(qoldiq=5800, min_z=MIN,
                        konteynerlar=[(55, s1["taklif"])])
        assert s2["taklif"] == 0


class TestChekkaHolatlar:
    def test_meyor_yoq(self):
        s = zanjir_sim(qoldiq=500, min_z=0, konteynerlar=[])
        assert s["xavf"] == "MEYOR_YOQ"
        assert s["taklif"] == 0

    def test_qoldiq_hozir_min_dan_past(self):
        s = zanjir_sim(qoldiq=1000, min_z=MIN, konteynerlar=[])
        assert s["uzilish_kun"] == 0
        assert s["xavf"] == "KRITIK"

    def test_taklif_50_ga_yaxlit(self):
        s = zanjir_sim(qoldiq=5830, min_z=MIN, konteynerlar=[])
        assert s["taklif"] % 50 == 0

    def test_konteyner_bugun_keladi(self):
        """kun_qoldi=0 konteyner darhol qo'shiladi."""
        s = zanjir_sim(qoldiq=1000, min_z=MIN, konteynerlar=[(0, 8000)])
        assert s["uzilish_kun"] is None
        assert s["taklif"] == 0

    def test_eski_kalitlar_saqlangan(self):
        """ui.py va kamomat_excel_v2 ishlatadigan kalitlar joyida."""
        s = zanjir_sim(qoldiq=5000, min_z=MIN, konteynerlar=[(20, 1000)])
        for k in ("uzilish_kun", "min_nuqta", "taklif_A", "taklif_B",
                  "taklif", "xavf"):
            assert k in s


class TestMaydaLimit:
    """2026-07-14: mayda truba (Ф<51) / profil (<50х50) — 200 limit."""

    def test_mayda_truba(self):
        from Generate_Asosiy_order import mayda_buyurtma_limiti as lim
        assert lim("Ф-19 ст 0,7 (6 м) (201 марка)", "Труба") == 200
        assert lim("Ф-38 ст 0,9 (6 м) (304 марка)", "Труба") == 200
        assert lim("Ф-51 ст 0,9 (6 м) (304 марка)", "Труба") == 0   # 51 mayda EMAS
        assert lim("Ф-76 ст 0,9 (6 м) (201 марка)", "Труба") == 0

    def test_mayda_profil(self):
        from Generate_Asosiy_order import mayda_buyurtma_limiti as lim
        assert lim("Пр. 20х20 ст 0,7 (6 м) (201 марка)", "Профиль") == 200
        assert lim("Пр. 40х20 ст 0,9 (6 м) (201 марка)", "Профиль") == 200
        assert lim("Пр. 50х50 ст 0,9 (6 м) (201 марка)", "Профиль") == 0  # 50х50 EMAS
        assert lim("Пр. 80х40 ст 1,1 (6 м) (201 марка)", "Профиль") == 0  # bir tomoni 50+

    def test_besshovniy_va_boshqalar_mustasno(self):
        from Generate_Asosiy_order import mayda_buyurtma_limiti as lim, get_category
        bes = "Ф-25 ст 3,0 Бесшовный (6 м) (304 марка)"
        assert get_category(bes) == "Безшовный труба"   # imlo (С) tuzatildi
        assert lim(bes, get_category(bes)) == 0          # limit yo'q
        assert get_category("Ф-32 ст 3,0 Б/Ш (6 м)") == "Безшовный труба"
        assert lim("Лист- 0,5 (1220х2440) (201 марка)", "Лист") == 0

    def test_truba_kategoriya_buzilmagan(self):
        from Generate_Asosiy_order import get_category
        assert get_category("Ф-51 ст 0,9 (6 м) (304 марка)") == "Труба"
