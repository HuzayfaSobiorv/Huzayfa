# -*- coding: utf-8 -*-
"""
Yuklama_optimal.py  — ABC asosida optimal konteyner yuklash.

12m konteyner: faqat 6.0 m Truба/Profil, jami 28 t, kamida 3 xil tovar
6m  konteyner: Truба/Profil (5.8 m) <=11 t  +  List <=18 t, jami <=28 t
Ustunlik: ABC sinf (A->B->C) -> Holat (KRITIK->PAST) -> Kam miqdori
"""
import re, sys, logging
from pathlib import Path

_THIS_DIR = Path(__file__).resolve().parent
if str(_THIS_DIR) not in sys.path:
    sys.path.insert(0, str(_THIS_DIR))

import pandas as pd
logger = logging.getLogger(__name__)

LIMIT_TOTAL        = 28_000
LIMIT_TRUBA_PROFIL = 11_000   # 6m konteyner uchun (backward compat)
LIMIT_LIST         = 18_000   # 6m konteyner uchun (backward compat)
LIMITS_BY_TYPE = {
    "12m": {"truba_profil": 28_000, "list":      0},
    "6m":  {"truba_profil": 11_000, "list": 18_000},
}
ABC_CAP_PCT = {"A": 0.40, "B": 0.60, "C": 1.00}
ABC_RANK    = {"A": 0, "B": 1, "C": 2}


def _holat_rank(holat: str) -> int:
    s = str(holat)
    if any(x in s for x in ("КРИТ", "CRIT")):   return 0
    if any(x in s for x in ("ПАСТ", "PAST")):   return 1
    if any(x in s for x in ("НОРМ", "NORM")):   return 2
    return 3


def get_length(name: str) -> float | None:
    m = re.search(r'\(([\d,\.]+)\s*[мm]\)', str(name))
    return float(m.group(1).strip().replace(",", ".")) if m else None


def get_category(name: str) -> str:
    s = str(name).strip()
    if re.match(r'^(\([^)]*\)\s*)?\u0424-\d+', s):   return "\u0422\u0440\u0443\u0431\u0430"
    if re.match(r'^(\([^)]*\)\s*)?\u041f\u0440\.\s*\d+', s): return "\u041f\u0440\u043e\u0444\u0438\u043b\u044c"
    if s.startswith("\u041b\u0438\u0441\u0442"):         return "\u041b\u0438\u0441\u0442"
    return "\u0411\u043e\u0448\u049b\u0430"


def yuk_turi(name: str, cat: str) -> str:
    if cat == "\u041b\u0438\u0441\u0442":
        return "6m"
    return "12m" if get_length(name) == 6.0 else "6m"


def _yangi_yuk(turi: str) -> dict:
    return {"turi": turi, "truba_profil_kg": 0.0,
            "list_kg": 0.0, "jami_kg": 0.0, "items": [], "_item_kg": {}}


def abc_map_yuklash() -> dict:
    try:
        import openpyxl
        f = _THIS_DIR / "abc_lookup.xlsx"
        if not f.exists():
            logger.warning("abc_lookup.xlsx topilmadi"); return {}
        wb = openpyxl.load_workbook(f, data_only=True, read_only=True)
        res = {}
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row and len(row) >= 3 and isinstance(row[1], str) and row[2] in ("A","B","C"):
                    res[row[1].strip()] = row[2]
        wb.close()
        logger.info("abc_map: %d ta tovar", len(res))
        return res
    except Exception as e:
        logger.error("abc_map_yuklash: %s", e); return {}


def tovar_vazni(tovar_nomi: str) -> float | None:
    try:
        from services import tovar_vazni_pb
        return tovar_vazni_pb(tovar_nomi)
    except Exception:
        return None


def optimallashtir(
    kerak_df: pd.DataFrame,
    mavjud_df: pd.DataFrame,
    abc_map: dict | None = None,
    max_yuklar: int = 20,
    xitoy_vazn: dict | None = None,
) -> tuple[list[dict], list[dict]]:
    """kerak_df: Товар|Холат|Кам  mavjud_df: Товар|Миқдор
    xitoy_vazn: {tovar_nomi: 1_dona_kg} — vazn_lookup da yo'q tovarlar uchun fallback
    """
    if abc_map is None:
        abc_map = abc_map_yuklash()
    mavjud = dict(zip(mavjud_df["Товар"], mavjud_df["Миқдор"]))
    kerak = kerak_df.copy()
    kerak["_ar"] = kerak["Товар"].map(lambda t: ABC_RANK.get(abc_map.get(str(t).strip(), "C"), 2))
    kerak["_hr"] = kerak["Холат"].map(_holat_rank)
    kerak = kerak.sort_values(["_ar", "_hr", "Кам"], ascending=[True, True, False])

    yuklar: list[dict] = []
    qolgan: list[dict] = []

    for _, row in kerak.iterrows():
        tovar      = row["Товар"]
        kerak_dona = int(row["Кам"])
        bor_dona   = int(mavjud.get(tovar, 0))
        if bor_dona <= 0:
            continue
        # Кам=0 bo'lsa (НОРМА/ПАСТ) — Xitoydan kelayotgan bo'lsa ham yuklanadi.
        # qoldi = bor_dona (barcha mavjud dona yuklanadi)
        vazn_dona = tovar_vazni(tovar)
        # Xitoy faylidan fallback: vazn_lookup da yo'q tovar uchun
        if (not vazn_dona or vazn_dona <= 0) and xitoy_vazn:
            vazn_dona = xitoy_vazn.get(str(tovar).strip())
        if not vazn_dona or vazn_dona <= 0:
            continue
        cat  = get_category(tovar)
        turi = yuk_turi(tovar, cat)
        if cat in ("\u0411\u043e\u0448\u049b\u0430",):
            continue
        key     = "list_kg"          if cat == "\u041b\u0438\u0441\u0442" else "truba_profil_kg"
        lim_key = "list"             if cat == "\u041b\u0438\u0441\u0442" else "truba_profil"
        abc_s   = abc_map.get(str(tovar).strip(), "C")
        cap_pct = ABC_CAP_PCT[abc_s]
        # Xitoyda nechta bo'lsa, hammasi yuklanadi (kerak formulasi faqat tartib uchun)
        qoldi   = bor_dona

        while qoldi > 0:
            for yuk in yuklar:
                if yuk["turi"] != turi:
                    continue
                slot_limit  = LIMITS_BY_TYPE[turi][lim_key]
                if slot_limit <= 0:
                    continue
                already_kg  = yuk["_item_kg"].get(tovar, 0.0)
                qolgan_cap  = slot_limit * cap_pct - already_kg
                qolgan_slot = slot_limit            - yuk[key]
                qolgan_jami = LIMIT_TOTAL           - yuk["jami_kg"]
                sig_kg      = min(qolgan_cap, qolgan_slot, qolgan_jami)
                sig_dona    = int(sig_kg // vazn_dona)
                if sig_dona <= 0:
                    continue
                dona = min(qoldi, sig_dona)
                og   = round(dona * vazn_dona, 2)
                yuk["items"].append({"tovar": tovar, "dona": dona, "vazn_kg": og, "abc": abc_s})
                yuk[key]               = round(yuk[key]  + og, 2)
                yuk["jami_kg"]         = round(yuk["jami_kg"] + og, 2)
                yuk["_item_kg"][tovar] = round(already_kg + og, 2)
                qoldi        -= dona
                mavjud[tovar] = mavjud.get(tovar, 0) - dona
                if qoldi == 0:
                    break
            if qoldi == 0:
                break
            if len(yuklar) >= max_yuklar:
                qolgan.append({"tovar": tovar, "dona": qoldi,
                               "vazn_kg": round(qoldi * vazn_dona, 2)})
                break
            yuklar.append(_yangi_yuk(turi))

    return yuklar, qolgan


def konteyner_xulosa(yuklar: list[dict]) -> dict:
    res = {"12m": 0, "6m": 0, "jami_kg": 0.0}
    for y in yuklar:
        res[y["turi"]] += 1
        res["jami_kg"] += y["jami_kg"]
    res["jami_t"] = round(res["jami_kg"] / 1000, 2)
    return res


def chop_natija(yuklar: list[dict], qolgan: list[dict] | None = None):
    for i, yuk in enumerate(yuklar, 1):
        print("=" * 68)
        print("KONTEYNER #%d  [%s]  --  Jami: %.0f kg  (TP: %.0f  List: %.0f)"
              % (i, yuk["turi"], yuk["jami_kg"],
                 yuk["truba_profil_kg"], yuk["list_kg"]))
        print("-" * 68)
        for it in yuk["items"]:
            print("  [%s] %-52s %5d d  %7.0f kg"
                  % (it["abc"], it["tovar"][:52], it["dona"], it["vazn_kg"]))
    if qolgan:
        print("=" * 68)
        print("QOLGAN (keyingi oyga): %d tovar" % len(qolgan))
        for it in qolgan:
            print("  %-52s %5d d  %7.0f kg"
                  % (it["tovar"][:52], it["dona"], it["vazn_kg"]))
    xs = konteyner_xulosa(yuklar)
    print("=" * 68)
    print("JAMI: %d konteyner  (12m: %d  6m: %d)  --  %.1f tonna"
          % (len(yuklar), xs["12m"], xs["6m"], xs["jami_t"]))


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s -- %(message)s")
    abc = abc_map_yuklash()
    print("ABC: %d ta  (A=%d B=%d C=%d)" % (
        len(abc), sum(v=="A" for v in abc.values()),
        sum(v=="B" for v in abc.values()),
        sum(v=="C" for v in abc.values())))
    tests = [
        "Ф-51 ст 0,9 (5,8 м) (201 марка)",
        "Ф-51 ст 0,9 (6 м) (201 марка)",
        "Пр. 30х30 ст 2,0 (6 м) (201 марка)",
    ]
    for t in tests:
        v = tovar_vazni(t)
        print(f"  {t}: {v} kg")
