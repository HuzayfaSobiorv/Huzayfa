# -*- coding: utf-8 -*-
"""
common.py sof funksiyalari uchun testlar.
Ishga tushirish (loyiha papkasida):  python -m pytest tests/ -v
"""
import json
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from common import (
    normalize_product_name,
    get_category,
    get_marka,
    get_category_with_marka,
    parse_qoldiq_str,
    yaxlitla_50,
    hisobla_min_zaxira,
    min_dan_kunlik_chiqar,
    atomic_json_write,
    KELISH_KUNI,
    get_himoya_foiz,
)


# ── normalize_product_name ──────────────────────────────────────────────
class TestNormalize:
    def test_oddiy_nom_ozgarmasin(self):
        assert normalize_product_name("Ф-51 ст 0,9 (6 м) (304 марка)") == \
            "Ф-51 ст 0,9 (6 м) (304 марка)"

    def test_ortiqcha_boshliqlar(self):
        assert normalize_product_name("  Ф-51   ст  0,9  (6 м)  ") == \
            "Ф-51 ст 0,9 (6 м)"

    def test_m_yopishib_yozilgan(self):
        # 2026-07-11 bug: "(6м)" bo'shliqsiz — inventar bilan mos kelmasdi
        assert "(6 м)" in normalize_product_name("Ф-51 ст 0,9 (6м) (304 марка)")

    def test_m_qoshiladi_raqamli_qavsga(self):
        # "(6)" → "(6 м)" — marka qavsiga tegilmaydi
        natija = normalize_product_name("Ф-51 ст 0,9 (6) (304 марка)")
        assert "(6 м)" in natija
        assert "(304 марка)" in natija

    def test_list_defis(self):
        assert normalize_product_name("Лист-0,8 (1220х2440)").startswith("Лист- 0,8")


# ── get_category ────────────────────────────────────────────────────────
class TestKategoriya:
    @pytest.mark.parametrize("nom,kutilgan", [
        ("Ф-51 ст 0,9 (6 м) (304 марка)", "ТРУБА"),
        ("Пр. 20х20 ст 0,9 (6 м) (304 марка)", "ПРОФИЛЬ"),
        ("Лист- 0,8 (1220х2440) (304 марка)", "ЛИСТ"),
        ("Лист рулон 0,5 (201 марка)", "ЛИСТ РУЛОН"),
        ("Баласина №25 (304 марка)", "БАЛАСИНА"),
        ("Стойка №30", "СТОЙКА"),
        ("Шар 51", "ШАР"),
        ("Номаълум tovar", "БОШҚА"),
    ])
    def test_kategoriyalar(self, nom, kutilgan):
        assert get_category(nom) == kutilgan

    def test_marka(self):
        assert get_marka("Ф-51 ст 0,9 (6 м) (304 марка)") == "304"
        assert get_marka("Стойка №30") == ""

    def test_kategoriya_marka_bilan(self):
        assert get_category_with_marka("Ф-51 ст 0,9 (6 м) (304 марка)") == "ТРУБА (304)"
        # Baласина'ga marka qo'shilmaydi
        assert get_category_with_marka("Баласина №25 (304 марка)") == "БАЛАСИНА"


# ── parse_qoldiq_str ────────────────────────────────────────────────────
class TestParseQoldiq:
    @pytest.mark.parametrize("kirish,kutilgan", [
        ("1 250", 1250),
        ("1,250", 1250),
        ("500/20", 500),      # '/' dan keyingisi tashlanadi
        ("-", 0),
        ("", 0),
        (None, 0),
        ("abc", 0),
    ])
    def test_qiymatlar(self, kirish, kutilgan):
        assert parse_qoldiq_str(kirish) == kutilgan


# ── min zaxira matematikasi ─────────────────────────────────────────────
class TestMinZaxira:
    def test_yaxlitla_50(self):
        assert yaxlitla_50(0) == 0
        assert yaxlitla_50(-5) == 0
        assert yaxlitla_50(1) == 50
        assert yaxlitla_50(50) == 50
        assert yaxlitla_50(51) == 100

    def test_min_va_kunlik_teskari(self):
        """hisobla_min_zaxira ↔ min_dan_kunlik_chiqar bir-biriga teskari bo'lishi kerak
        (50 ga yaxlitlash farqi doirasida)."""
        kunlik = 10.0
        kat = "ТРУБА"
        min_z = hisobla_min_zaxira(kunlik, kat)
        qayta = min_dan_kunlik_chiqar(min_z, kat)
        # yaxlitlash tufayli qayta >= kunlik, lekin 50/[denom] dan ko'p farq qilmasin
        farq_limiti = 50 / (KELISH_KUNI * (1 + get_himoya_foiz(kat)))
        assert kunlik <= qayta <= kunlik + farq_limiti + 0.001


# ── atomic_json_write ───────────────────────────────────────────────────
class TestAtomicWrite:
    def test_yozadi_va_oqiladi(self, tmp_path):
        p = tmp_path / "holat.json"
        data = {"kanal": "asosiy", "tovarlar": {"Ф-51": 100}}
        atomic_json_write(p, data, indent=2)
        assert json.loads(p.read_text(encoding="utf-8")) == data

    def test_eski_fayl_ustiga_yozadi(self, tmp_path):
        p = tmp_path / "holat.json"
        p.write_text('{"eski": true}', encoding="utf-8")
        atomic_json_write(p, {"yangi": 1})
        assert json.loads(p.read_text(encoding="utf-8")) == {"yangi": 1}

    def test_tmp_qoldiq_qolmaydi(self, tmp_path):
        p = tmp_path / "holat.json"
        atomic_json_write(p, [1, 2, 3])
        qoldiqlar = [f for f in os.listdir(tmp_path) if ".tmp" in f]
        assert qoldiqlar == []

    def test_xato_bolsa_eski_fayl_buzilmaydi(self, tmp_path):
        p = tmp_path / "holat.json"
        atomic_json_write(p, {"muhim": "malumot"})
        # JSON ga aylanmaydigan obyekt — xato beradi
        with pytest.raises(TypeError):
            atomic_json_write(p, {"x": object()})
        # eski fayl butunligicha qoladi
        assert json.loads(p.read_text(encoding="utf-8")) == {"muhim": "malumot"}
