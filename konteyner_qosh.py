"""
konteyner_qosh.py — Xitoy装箱单 (Truba/Profil) va出货清单 (List) parser
Konteynerlar yo'lda ro'yxatini shakllantiradi.

Parser ikki faylni o'qib, ISO konteyner raqami bo'yicha birlashtiradi.
Chiqish: [{"iso": "TCKU2238508", "sana": "21.03.2026", "items": [("Tovar nomi", 100), ...], "manba": "aralash"}]
"""
import re
import math
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import openpyxl

# ISO konteyner raqami: 4 harf + 7 raqam
_ISO_RE = re.compile(r'\b([A-Z]{4}\d{7})\b')

# ── Yordamchi funksiyalar ─────────────────────────────────────────────────────

def _iso(text) -> str | None:
    """ISO konteyner raqamini matndan ajratib oladi."""
    m = _ISO_RE.search(str(text))
    return m.group(1) if m else None


def _sana_format(d) -> str:
    """datetime → '21.03.2026'"""
    if isinstance(d, (datetime, date)):
        return d.strftime("%d.%m.%Y") if isinstance(d, datetime) else d.strftime("%d.%m.%Y")
    return str(d)


def _yaxlitla_stenka(s: str) -> str:
    """Xitoy stenka → inventar stenka (0.85 → 0,9 va h.k.)"""
    try:
        v = float(str(s).replace(',', '.'))
    except (ValueError, TypeError):
        return str(s).replace('.', ',')
    TABLE = [
        (0.50, '0,5'), (0.55, '0,6'), (0.60, '0,6'), (0.63, '0,65'),
        (0.65, '0,65'), (0.68, '0,7'), (0.70, '0,7'), (0.75, '0,8'),
        (0.80, '0,8'), (0.85, '0,9'), (0.90, '0,9'), (0.95, '1,0'),
        (1.00, '1,0'), (1.05, '1,1'), (1.10, '1,1'), (1.20, '1,2'),
        (1.25, '1,3'), (1.30, '1,3'), (1.35, '1,35'), (1.40, '1,4'),
        (1.45, '1,45'), (1.50, '1,5'), (1.75, '1,75'), (2.00, '2,0'),
        (2.50, '2,5'), (3.00, '3,0'),
    ]
    best = min(TABLE, key=lambda x: abs(x[0] - v))
    return best[1]


def _truba_spec_to_name(spec: str, marka_raw) -> str | None:
    """
    Truba/Profil Xitoy speci → inventar nomi.
    '0.85*50.8*5800' + 201 → 'Ф-51 ст 0,9 (6 м) (201 марка)'
    """
    spec = str(spec).strip()
    # Marka
    marka = str(int(float(marka_raw))) if isinstance(marka_raw, float) else str(marka_raw).strip()
    if marka not in ('201', '304', '316', '321', '430'):
        marka = '201'

    # Format: stenka * diametr * uzunlik  (mm)
    m = re.match(r'^([\d\.]+)[*×x]([\d\.]+)[*×x]([\d\.]+)$', spec)
    if not m:
        return None

    try:
        stenka_raw = float(m.group(1))
        dim2       = float(m.group(2))
        uzunlik_mm = float(m.group(3))
    except ValueError:
        return None

    stenka = _yaxlitla_stenka(stenka_raw)
    uzunlik_m = uzunlik_mm / 1000.0
    if uzunlik_m == int(uzunlik_m):
        uz_s = str(int(uzunlik_m))
    else:
        uz_s = f"{uzunlik_m:.1f}".replace('.', ',')
    uz_str = f"{uz_s} м"
    marka_sfx = f" ({marka} марка)"

    # Труба: stenka * diametr (yumaloq)  → diametr ~ 50.8 → Ф-51
    if dim2 <= 200 and dim2 != int(dim2):
        # Yumaloq truba: diametr yaxlitlanadi
        d_int = round(dim2)
        return f"Ф-{d_int} ст {stenka} ({uz_str}){marka_sfx}"

    # Профиль: stenka * a * b yoki stenka * a (kvadrat)
    # Lekin bizning formatda: stenka*kenglik*uzunlik
    # Профиль kvadrat: dim2 = a = b
    if dim2 == int(dim2):
        a = int(dim2)
        return f"Пр. {a}х{a} ст {stenka} ({uz_str}){marka_sfx}"

    return None


def _list_spec_to_name(spec: str, marka_raw, rang: str = '') -> str | None:
    """
    List Xitoy speci → inventar nomi.
    '0.75*1219*2438' + 201 → 'Лист 0,75 (1219х2438) (201 марка)'
    """
    spec = str(spec).strip()
    marka = str(int(float(marka_raw))) if isinstance(marka_raw, float) else str(marka_raw).strip()
    if marka not in ('201', '304', '316', '321', '430'):
        marka = '201'

    m = re.match(r'^([\d\.]+)[*×x]([\d\.]+)[*×x]([\d\.]+)$', spec)
    if not m:
        return None

    try:
        qalinlik = float(m.group(1))
        en       = int(float(m.group(2)))
        boy      = int(float(m.group(3)))
    except ValueError:
        return None

    # Qalinlik formati: 0.75 → '0,75', 1.0 → '1,0'
    if qalinlik == int(qalinlik):
        q_s = f"{int(qalinlik)},0"
    else:
        q_s = f"{qalinlik:.2f}".rstrip('0').replace('.', ',')

    # Rang prefiksi
    rang_map = {
        '砂板': '', '8K钛金': '(Голд) ', '精磨8K': '', '8K黑钛': '(Кора) ', '8K': '',
    }
    pfx = rang_map.get(rang, '')

    return f"{pfx}Лист {q_s} ({en}х{boy}) ({marka} марка)"


# ── Truba/Profil parser (装箱单) ──────────────────────────────────────────────

def _parse_truba_zhuangxiang(raw: bytes) -> dict:
    """
    装箱单 (Truba/Profil packing list) ni o'qiydi.
    Qaytaradi: {iso_no: {"sana": date, "items": [(tovar, miqdor), ...]}}
    """
    wb = openpyxl.load_workbook(BytesIO(raw), data_only=True)
    ws = wb.active

    # 1) Merged celllardan 柜号 → qator_raqami
    kont_at = {}  # row_idx → iso_no
    for mc in ws.merged_cells.ranges:
        val = ws.cell(mc.min_row, mc.min_col).value
        if val and '柜号' in str(val):
            iso = _iso(str(val))
            if iso:
                kont_at[mc.min_row] = iso

    # 2) Blok boshlarini topish (B ustun '柜' bilan boshlanadi)
    blocks = []  # (start_row, gui_name, sana)
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        b = row[1].value
        if b and str(b).strip().startswith('柜'):
            sana = row[8].value if len(row) > 8 else None
            if isinstance(sana, datetime):
                sana = sana.date()
            blocks.append((row[0].row, str(b).strip(), sana))

    result = {}

    for i, (start_row, gui, sana) in enumerate(blocks):
        end_row = blocks[i + 1][0] if i + 1 < len(blocks) else ws.max_row + 1

        # Shu blok ichidagi ISO raqam
        iso = None
        for kr in sorted(kont_at):
            if start_row <= kr < end_row:
                iso = kont_at[kr]
                break
        if not iso:
            continue

        # Mahsulotlarni o'qish
        items = []
        in_products = False
        for ri in range(start_row, end_row):
            row_vals = [ws.cell(ri, c).value for c in range(1, 10)]
            b_val = row_vals[1]  # B ustun (0-indexed: index 1 = column B)
            # Sarlavha satrini aniqlash
            if b_val == '规格':
                in_products = True
                continue
            if not in_products:
                continue
            spec   = row_vals[1]   # B = 规格
            marka  = row_vals[2]   # C = 材质
            # zhishu = row_vals[4]  # E = 支数 (total piece count)
            zhishu = row_vals[4]   # index 4 = column F? Let me recheck
            if not spec or not marka:
                continue
            spec_s = str(spec).strip()
            if not re.match(r'^[\d\.]+[*×x]', spec_s):
                continue
            try:
                miqdor = int(float(zhishu)) if zhishu else 0
            except (ValueError, TypeError):
                miqdor = 0
            if miqdor <= 0:
                continue
            nom = _truba_spec_to_name(spec_s, marka)
            if nom:
                items.append((nom, miqdor))
            else:
                items.append((spec_s, miqdor))

        if iso not in result:
            result[iso] = {"sana": sana, "items": []}
        result[iso]["items"].extend(items)

    return result


# ── List parser (出货清单) ────────────────────────────────────────────────────

def _parse_list_chuhuo(raw: bytes) -> dict:
    """
    出货清单 (List packing list) ni o'qiydi.
    Qaytaradi: {iso_no: {"sana": date, "items": [(tovar, miqdor), ...]}}
    """
    try:
        import xlrd
    except ImportError:
        raise ImportError("xlrd o'rnatilmagan: pip install xlrd")

    book = xlrd.open_workbook(file_contents=raw)
    sh   = book.sheets()[0]

    # Hujjat sanasi (row 2, col 1)
    doc_sana = None
    if sh.nrows > 1:
        v = sh.row_values(1)[1]
        if isinstance(v, float):
            try:
                doc_sana = xlrd.xldate_as_datetime(v, book.datemode).date()
            except Exception:
                pass

    # Sarlavha satrini topish
    hdr_row = 2
    for ri in range(min(5, sh.nrows)):
        row = sh.row_values(ri)
        if '规格' in row or any('规格' in str(c) for c in row):
            hdr_row = ri
            break

    # Ustun indekslari
    hdrs = [str(c).strip() for c in sh.row_values(hdr_row)]
    mat_i  = next((i for i, h in enumerate(hdrs) if '材质' in h), 0)
    rang_i = next((i for i, h in enumerate(hdrs) if '颜色' in h), 1)
    gg_i   = next((i for i, h in enumerate(hdrs) if '规格' in h), 2)
    qty_i  = next((i for i, h in enumerate(hdrs) if '数量' in h), 3)
    bz_i   = next((i for i, h in enumerate(hdrs) if '备注' in h), 9)

    result   = {}
    cur_iso  = None

    for ri in range(hdr_row + 1, sh.nrows):
        row = sh.row_values(ri)
        if len(row) <= bz_i:
            continue

        # Yangi konteyner bloki?
        note = str(row[bz_i]).strip() if row[bz_i] else ''
        iso  = _iso(note)
        if iso:
            cur_iso = iso
            if cur_iso not in result:
                result[cur_iso] = {"sana": doc_sana, "items": []}

        if not cur_iso:
            continue

        spec  = str(row[gg_i]).strip() if row[gg_i] else ''
        marka = row[mat_i]
        rang  = str(row[rang_i]).strip() if rang_i < len(row) and row[rang_i] else ''

        if not re.match(r'^[\d\.]+[*×x]', spec):
            continue
        if any(k in spec for k in ('合计', '小计', '序号')):
            continue

        try:
            miqdor = int(float(row[qty_i])) if qty_i < len(row) and row[qty_i] else 0
        except (ValueError, TypeError):
            miqdor = 0
        if miqdor <= 0:
            continue

        nom = _list_spec_to_name(spec, marka, rang)
        if nom:
            result[cur_iso]["items"].append((nom, miqdor))
        else:
            result[cur_iso]["items"].append((spec, miqdor))

    return result


# ── Asosiy birlashtirish funksiyasi ──────────────────────────────────────────

def xitoy_yuklar_oqi(truba_raw: bytes, list_raw: bytes) -> list[dict]:
    """
    Ikki faylni o'qib, konteynerlar ro'yxatini qaytaradi.
    Qaytaradi: [
      {
        "iso":   "TCKU2238508",
        "sana":  "21.03.2026",
        "items": [("Ф-51 ст 0,9 (6 м) (201 марка)", 100), ...],
        "manba": "aralash" | "truba" | "list"
      }, ...
    ]
    """
    truba_map = _parse_truba_zhuangxiang(truba_raw)
    list_map  = _parse_list_chuhuo(list_raw)

    all_isos = set(truba_map) | set(list_map)
    result   = []

    for iso in sorted(all_isos):
        t_data = truba_map.get(iso)
        l_data = list_map.get(iso)

        if t_data and l_data:
            manba = "aralash"
            sana  = t_data["sana"] or l_data["sana"]
            items = t_data["items"] + l_data["items"]
        elif t_data:
            manba = "truba"
            sana  = t_data["sana"]
            items = t_data["items"]
        else:
            manba = "list"
            sana  = l_data["sana"]
            items = l_data["items"]

        sana_s = _sana_format(sana) if sana else "?"

        result.append({
            "iso":   iso,
            "sana":  sana_s,
            "items": items,
            "manba": manba,
        })

    return result


# ── Mavjud konteynerlar bilan solishtirish ────────────────────────────────────

def yangi_konteynerlar(yuklar: list[dict], kont_dir: Path) -> list[dict]:
    """
    Allaqachon saqlangan xlsx fayllar bilan solishtiradi.
    Faqat yangi (mavjud bo'lmagan) konteynerlarni qaytaradi.
    """
    existing = set()
    if kont_dir.exists():
        for f in kont_dir.glob("*.xlsx"):
            # CRXU1561318_07.06.2026.xlsx yoki _D.xlsx
            iso = f.stem.split('_')[0]
            existing.add(iso)

    return [k for k in yuklar if k["iso"] not in existing]


# ── Konteyner xlsx yaratish ───────────────────────────────────────────────────

def konteyner_xlsx_yarat(kont: dict, kont_dir: Path) -> Path:
    """
    Bitta konteyner uchun xlsx fayl yaratadi.
    Fayl nomi: TCKU2238508_21.03.2026.xlsx
    Ustunlar: Товар | Миқдор
    """
    kont_dir.mkdir(parents=True, exist_ok=True)
    sana_fayl = kont["sana"].replace('.', '-') if kont["sana"] != "?" else "nomalum"
    # Fayl nomi: ISO_DD-MM-YYYY.xlsx
    fname = f"{kont['iso']}_{kont['sana']}.xlsx"
    fpath = kont_dir / fname

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Yuk"
    ws.append(["Товар", "Миқдор"])
    for tovar, miqdor in kont["items"]:
        ws.append([tovar, miqdor])

    # Ustun kengligi
    ws.column_dimensions['A'].width = 55
    ws.column_dimensions['B'].width = 12

    wb.save(fpath)
    return fpath


# ── Preview matn ─────────────────────────────────────────────────────────────

def preview_matn(yangilar: list[dict]) -> str:
    """
    Foydalanuvchiga ko'rsatiladigan qisqacha xabar.
    """
    if not yangilar:
        return "✅ Barcha konteynerlar allaqachon ro'yxatda."

    aralash = [k for k in yangilar if k["manba"] == "aralash"]
    faqat_t = [k for k in yangilar if k["manba"] == "truba"]
    faqat_l = [k for k in yangilar if k["manba"] == "list"]

    lines = [f"🆕 *{len(yangilar)} ta yangi konteyner topildi:*\n"]

    if aralash:
        lines.append("🔀 *Birlashtirilgan (Трубa + Лист):*")
        for k in aralash:
            lines.append(f"  • `{k['iso']}` — {k['sana']} — {len(k['items'])} ta tovar")

    if faqat_t:
        lines.append("\n🔩 *Faqat Труба/Профиль:*")
        for k in faqat_t:
            lines.append(f"  • `{k['iso']}` — {k['sana']} — {len(k['items'])} ta tovar")

    if faqat_l:
        lines.append("\n📄 *Faqat Лист:*")
        for k in faqat_l:
            lines.append(f"  • `{k['iso']}` — {k['sana']} — {len(k['items'])} ta tovar")

    lines.append("\n✅ Tasdiqlaysizmi?")
    return "\n".join(lines)
