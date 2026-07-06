"""
config.py — NEJAVIYKA Bot sozlamalari
Barcha global konstant, path va konfiguratsiya shu yerda.
"""
import logging, os, threading, time
from pathlib import Path

# ── Papkalar ────────────────────────────────────────────────────────────────
BASE_DIR      = Path(__file__).resolve().parent
DATA_FILE     = BASE_DIR / "chiqish" / "NEJAVIYKA_POWER_BI.xlsx"
BOT_HOLAT_DIR = BASE_DIR / "bot_holat"
BOT_HOLAT_DIR.mkdir(exist_ok=True)
KONT_DIR      = BOT_HOLAT_DIR / "konteynerlar"
KONT_DIR.mkdir(exist_ok=True)
XITOY_PARSED_DIR = BASE_DIR / "konteynerlar" / "xitoy_parsed"
# Bir marta tasdiqlangan konteyner ISO'lari tarixi — fayli keyinchalik
# o'chirilsa ham (masalan KELDI bo'lib arxivlangach), qayta qo'shilib
# ketmasligi uchun (konteyner_qosh.py / handlers.py "yo'lga qo'shish" flow)
KONTEYNER_TARIX_FILE = BOT_HOLAT_DIR / "qoshilgan_konteynerlar.json"
# Admin tomonidan bir marta qo'lda to'g'irlangan Xitoy spec -> inventar nomi
# bog'lanishlari — shu yerga saqlanadi, keyingi safar xuddi shu spec
# chiqqanda avtomatik (qayta taxmin qilmasdan) ishlatiladi.
XITOY_TUZATISH_FILE = BOT_HOLAT_DIR / "xitoy_tuzatishlar.json"
ENV_FILE      = BASE_DIR / ".env"

# ── .env fayl yuklash ────────────────────────────────────────────────────────
def _load_env(path: Path = ENV_FILE) -> None:
    if not path.exists():
        return
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))

_load_env()
BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()

# Super admin (1 ta) — foydalanuvchilarni boshqaradi, kirish xabarlari oladi
_super_raw = os.getenv("SUPER_ADMIN_ID", "").strip()
SUPER_ADMIN_ID: int | None = int(_super_raw) if _super_raw.isdigit() else None

# Admin IDlar: ADMIN_IDS=123,456,789 (vergul bilan, .env da)
# Super admin ham shu ro'yxatga kiradi
_admin_raw = os.getenv("ADMIN_IDS", os.getenv("ADMIN_ID", "")).strip()
ADMIN_IDS: set[int] = set(
    int(x.strip()) for x in _admin_raw.split(",") if x.strip().isdigit()
)
if SUPER_ADMIN_ID:
    ADMIN_IDS.add(SUPER_ADMIN_ID)
ADMIN_ID = SUPER_ADMIN_ID or next(iter(ADMIN_IDS), None)  # backward compat

# Filyal uchun bog'lanish ma'lumotlari
SUPPORT_PHONE    = os.getenv("SUPPORT_PHONE", "").strip()
SUPPORT_USERNAME = os.getenv("SUPPORT_USERNAME", "").strip()

# "Kelgan yuklar" guruhi — konteyner KELDI bo'lganda avtomatik rasm shu yerga boradi
_kg_chat_raw  = os.getenv("KELGAN_YUKLAR_CHAT_ID", "").strip()
_kg_topic_raw = os.getenv("KELGAN_YUKLAR_TOPIC_ID", "").strip()
KELGAN_YUKLAR_CHAT_ID:  int | None = int(_kg_chat_raw)  if _kg_chat_raw.lstrip("-").isdigit()  else None
KELGAN_YUKLAR_TOPIC_ID: int | None = int(_kg_topic_raw) if _kg_topic_raw.lstrip("-").isdigit() else None

# ── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    format="%(asctime)s — %(name)s — %(levelname)s — %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("nejaviyka")

# ── Excel RAM cache (har 5 daqiqada yangilanadi) ─────────────────────────────
import pandas as pd

_XLSX_CACHE: dict = {"inv": None, "kont": None, "ts": 0.0}
_XLSX_LOCK  = threading.Lock()
_MPL_LOCK   = threading.Lock()   # matplotlib thread-safe emas
_CACHE_TTL  = 300                # sekund


def xlsx_refresh(force: bool = False) -> None:
    """Excel ni diskdan o'qib cache ga yozadi (faqat TTL o'tganda yoki force=True)."""
    now = time.monotonic()
    with _XLSX_LOCK:
        if not force and _XLSX_CACHE["inv"] is not None and (now - _XLSX_CACHE["ts"]) < _CACHE_TTL:
            return
        try:
            _XLSX_CACHE["inv"]  = pd.read_excel(DATA_FILE, sheet_name="Инвентар")
            _XLSX_CACHE["kont"] = pd.read_excel(DATA_FILE, sheet_name="Контейнерлар")
            _XLSX_CACHE["ts"]   = time.monotonic()
        except Exception as e:
            logger.warning(f"xlsx cache refresh xato: {e}")


def get_inv() -> "pd.DataFrame":
    xlsx_refresh()
    return _XLSX_CACHE["inv"]


def get_kont() -> "pd.DataFrame":
    xlsx_refresh()
    return _XLSX_CACHE["kont"]


# ── Mahsulot varaqlari ───────────────────────────────────────────────────────
VARAQLAR = ["Труба", "Профиль", "Лист", "Баласина", "Стойка", "Аксессуар"]

CAT_SHEET = {
    "ТРУБА": "Труба", "ПРОФИЛЬ": "Профиль",
    "ЛИСТ": "Лист",  "ЛИСТ РУЛОН": "Лист",
    "БАЛАСИНА": "Баласина", "СТОЙКА": "Стойка",
}
AKSESSUAR_KATS = {"ШАР", "ОТВОД", "СОККА", "ЧАШКА", "СОВУН", "КУЗИКОРИН", "БОШКА"}

CH_KEY = {"asosiy": "ch_asosiy", "sex": "ch_sex", " osh": "ch_osh"}

# Navigatsiya: qaysi ekrandan "Orqaga" qayerga boradi
BACK_MAP = {
    "order":         "main",
    "order_channel": "order",
    "load":          "main",
    "load_channel":  "load",
    "status":        "main",
    "settings":      "main",
    "search":        "main",
    "search_kat":    "search",
    "konteyner":     "main",
    "keldi_menu":    "konteyner",
    "keldi_ekran":   "keldi_menu",
}

# Xitoy nomi → inventar nomi qo'lda mapping
XITOY_NOM_MAP: dict[str, str] = {}
