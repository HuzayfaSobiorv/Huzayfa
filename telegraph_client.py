"""
telegraph_client.py — Telegra.ph API integratsiyasi
====================================================
- Token bir marta yaratiladi, .telegraph_token faylida saqlanadi
- PNG rasm telegra.ph/upload ga yuklanadi → URL qaytaradi
- create_karta_page → to'liq sahifa yaratadi, URL qaytaradi
"""

import json
import logging
import requests
from pathlib import Path
from datetime import datetime

logger = logging.getLogger(__name__)

TELEGRAPH_API    = "https://api.telegra.ph"
TELEGRAPH_UPLOAD = "https://telegra.ph/upload"
TOKEN_FILE       = Path(__file__).parent / ".telegraph_token"
REQUEST_TIMEOUT  = 30


# ============================================================
# TOKEN
# ============================================================
def get_token() -> str:
    """Token fayldan olinadi yoki yangi yaratiladi."""
    if TOKEN_FILE.exists():
        t = TOKEN_FILE.read_text(encoding="utf-8").strip()
        if t:
            return t
    token = _create_account()
    TOKEN_FILE.write_text(token, encoding="utf-8")
    return token


def _create_account() -> str:
    resp = requests.post(
        f"{TELEGRAPH_API}/createAccount",
        json={
            "short_name":  "NejaviykaBot",
            "author_name": "Nejaviyka Zaxira"
        },
        timeout=REQUEST_TIMEOUT
    )
    data = resp.json()
    if not data.get("ok"):
        raise RuntimeError(f"Telegraph createAccount xato: {data}")
    return data["result"]["access_token"]


# ============================================================
# RASM YUKLASH
# ============================================================
def upload_image(png_bytes: bytes) -> str:
    """PNG → telegra.ph ga yuklaydi, to'liq URL qaytaradi."""
    resp = requests.post(
        TELEGRAPH_UPLOAD,
        files={"file": ("chart.png", png_bytes, "image/png")},
        timeout=REQUEST_TIMEOUT
    )
    data = resp.json()
    # Javob: [{"src": "/file/abc123.png"}]
    if isinstance(data, list) and data:
        return "https://telegra.ph" + data[0]["src"]
    raise RuntimeError(f"Telegraph upload xato: {data}")


# ============================================================
# SAHIFA YARATISH
# ============================================================
def create_karta_page(
    token:     str,
    tovar:     str,
    holat:     str,
    qoldiq:    float,
    min_z:     float,
    yolda_j:   float,
    kont_rows: list,
    img_url:   str | None
) -> str:
    """
    To'liq tovar kartasi sahifasini yaratadi.
    kont_rows: [{"nom", "miqdor", "yukl", "kelish", "kun"}, ...]
    Qaytaradi: sahifa URL
    """
    nodes = []

    # ── Holat satri ────────────────────────────────────────────
    holat_txt = holat.replace("🔴", "🔴").replace("🟡", "🟡").replace("🟢", "🟢")
    nodes.append({"tag": "p", "children": [holat_txt]})

    # ── Asosiy ko'rsatkichlar ───────────────────────────────────
    nodes.append({"tag": "p", "children": [
        f"Qoldiq: {int(qoldiq):,}  |  Min zaxira: {int(min_z):,}  |  Yo'lda jami: {int(yolda_j):,}"
        .replace(",", " ")
    ]})

    # ── Rasm ───────────────────────────────────────────────────
    if img_url:
        nodes.append({"tag": "img", "attrs": {"src": img_url}})

    # ── Konteynerlar jadvali ────────────────────────────────────
    if kont_rows:
        nodes.append({"tag": "h4", "children": ["Yo'ldagi konteynerlar"]})

        # Sanalar bo'yicha guruhlash
        by_date: dict[str, list] = {}
        for kr in kont_rows:
            d = str(kr.get("kelish", "—"))
            by_date.setdefault(d, []).append(kr)

        for date_key in sorted(by_date.keys()):
            rows = by_date[date_key]
            kun  = rows[0].get("kun", 0)
            jami = sum(r["miqdor"] for r in rows)

            # Sana sarlavhasi
            kun_txt = f"{kun} kun" if kun > 0 else "bugun/kechikdi"
            nodes.append({"tag": "p", "children": [
                {"tag": "b", "children": [f"📅 {date_key}  ({kun_txt})  +{jami:,}".replace(",", " ")]},
            ]})

            # Har konteyner qatori
            for kr in rows:
                nodes.append({"tag": "p", "children": [
                    f"  📦 {kr['nom']}  —  {kr['miqdor']:,} dona  |  Yuklangan: {kr['yukl']}"
                    .replace(",", " ")
                ]})

    else:
        nodes.append({"tag": "p", "children": ["🚢 Yo'lda konteyner yo'q"]})

    # ── Yangilangan vaqt ────────────────────────────────────────
    now_str = datetime.now().strftime("%d.%m.%Y %H:%M")
    nodes.append({"tag": "p", "children": [f"🕐 {now_str} da yangilangan"]})

    # ── Sahifa yaratish ─────────────────────────────────────────
    title = tovar[:256] if tovar else "Tovar kartasi"
    resp = requests.post(
        f"{TELEGRAPH_API}/createPage",
        json={
            "access_token":   token,
            "title":          title,
            "content":        json.dumps(nodes),
            "return_content": False
        },
        timeout=REQUEST_TIMEOUT
    )
    data = resp.json()
    if data.get("ok"):
        return data["result"]["url"]
    raise RuntimeError(f"Telegraph createPage xato: {data}")
