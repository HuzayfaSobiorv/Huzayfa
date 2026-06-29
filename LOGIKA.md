# NERJAVEYKA BOT — Logika hujjati
_Oxirgi yangilanish: 2026-06-23_

---

## 1. BUYURTMA BERISH LOGIKASI

### Flow (qadamlar)

```
Asosiy menyu
  └─ 🛒 Buyurtma yig'ish
        └─ Kanal tanlash: Asosiy | CEX | Osh
              └─ order_channel ekrani:
                    ├─ 📋 Kamomadni ko'rish   → kamomat_ko'rish()
                    ├─ 📄 Buyurtma Excel olish → excel action
                    └─ ✅ Tasdiqlash           → tasdiq action
```

### 1.1 Buyurtma Excel olish (excel action)

1. Xitoy ostatka JSON tekshiriladi (`xitoy_yuklash(kanal)`)
2. **Agar JSON mavjud** → inline klaviatura:
   - `Mavjudni ishlatsin` — eski JSON bilan
   - `Yangi yuklash` — `kutilmoqda = ("xitoy_ostatka_fayl", kanal)` holatiga
   - `Hisobsiz ber` — xitoy_map={} bilan to'g'ridan Excel
3. **Agar JSON yo'q** → inline klaviatura:
   - `Xitoy faylini yuklang` — `kutilmoqda = ("xitoy_ostatka_fayl", kanal)`
   - `Hisobsiz ber` — xitoy_map={} bilan to'g'ridan Excel

4. Qaror kelingach → `draft_buyurtma_yubor(msg, context, kanal, lang, xitoy_ostatka=xitoy_map)`:
   - `asosiy_styled_excel_yarat(kanal, lang, xitoy_ostatka)` chaqiriladi
   - Tasdiqlangan buyurtmalar (`buyurtma_asosiy.json`) AYIRILADI
   - Xitoy ostatka K ustuniga qo'shiladi (yetarli bo'lsa buyurtma nol bo'ladi)
   - Excel fayl yaratiladi, foydalanuvchiga jo'natiladi
   - `draft_asosiy.json` saqlanadi (buyurtma_tekshir uchun lookup)

### 1.2 Tasdiqlash (tasdiq action)

1. `kutilmoqda = ("buyurtma_tasdiq", kanal)`
2. Foydalanuvchi avval olingan Excel faylini qaytarib yuboradi
3. `buyurtma_tekshir(fayl_bytes)`:
   - `draft_asosiy.json` dan lookup (inventar o'zgarsa ham ishlaydi)
   - Har bir varaq (Труба/Профиль/Лист/Баласина/Стойка/Аксессуар) o'qiladi
   - "Tovar nomi" sarlavhasi 1-5 qatorda qidiriladi
   - Buyurtma > 0 bo'lgan qatorlar `items` ro'yxatiga qo'shiladi
4. **Preview** ko'rsatiladi:
   - Jami tovar soni, varaq bo'yicha taqsimot
   - Inline: `✅ Tasdiqlash` | `❌ Bekor qilish`
5. **Tasdiqlaganda** → `buyurtma_saqlash(kanal, items)`:
   - `bot_holat/buyurtma_{kanal}.json` ga yoziladi
   - `sana` (bugungi sana) qo'shiladi
6. **Keyingi Excel olishda** tasdiqlangan buyurtmalar avtomatik ayiriladi

### 1.3 Xitoy ostatka (`xitoy_ostatka_fayl` holati)

- Fayl kelsa → `xitoy_ostatka_oqi()` → JSON ga qo'shib saqlanadi
- Ko'p fayl yuborish mumkin (akkumulyatsiya)
- JSON saqlanadi: `bot_holat/xitoy_{kanal}.json`
  - `tovarlar` — zakaz (K ustun, buyurtma hisoblash uchun)
  - `ombor` — tayyor (L ustun, yuklatish rejasi uchun)
- Tasdiqlangan buyurtma (`buyurtma_{kanal}.json`) avtomatik tozalanadi

---

## 2. KONTEYNER YUKLASH LOGIKASI

### Flow (qadamlar)

```
Asosiy menyu
  └─ 🚛 Konteyner yuklash
        └─ Kanal tanlash: Asosiy | CEX | Osh
              └─ load_channel ekrani:
                    └─ ▶️ Hisoblash tugmasi
```

### 2.1 load_channel ekraniga o'tish

Kanal tanlash → `go_screen(msg, context, "load_channel", kanal=kanal)` →
  - `kutilmoqda = ("xitoy_fayl", kanal)` — fayl kutish holati
  - `xitoy_akkum = {}` — akkumulyator reset
  - Ekranda: "Xitoy tayyor ro'yxatini yuboring" matni + `▶️ Hisoblash` + Orqaga

### 2.2 Fayl yuborish (`xitoy_fayl` holati)

Fayl kelsa → `fayl_keldi()` → `kut[0] == "xitoy_fayl"`:
1. `xitoy_ostatka_oqi(bytes)` → `ombor_map` (tayyor tovarlar, L ustun)
2. `xitoy_akkum[tovar] += miqdor` (bir nechta fayl qo'llab-quvvatlash)
3. "✅ N ta tovar qo'shildi. Jami: M ta." xabari
4. Yana fayl yuborsangiz — akkumulyatorga qo'shiladi
5. `kutilmoqda` o'zgarishsiz qoladi ("xitoy_fayl")

### 2.3 Hisoblash (hisoblash action)

`▶️ Hisoblash` bosilsa → `action == "hisoblash"`:
1. `akkum = context.user_data.get("xitoy_akkum", {})` tekshiriladi
2. Bo'sh bo'lsa → "⚠️ Hali hech qanday fayl yuborilmadi"
3. `yuklash_animatsiya` bilan `main_with_data(kanal, akkum)` chaqiriladi
4. `yuklatish_rejasi.py` → `Yuklatish_Rejasi_YYYY-MM-DD.xlsx` yaratiladi
5. Excel jo'natiladi caption bilan
6. `xitoy_akkum` va `kutilmoqda` tozalanadi
7. **Keyin**: `go_screen("load_channel", kanal)` — ekran yangilanadi,
   foydalanuvchi boshidan kanal tanlamay qayta hisoblash imkoni

### 2.4 yuklatish_rejasi.py qoidalari

- Jami yuk ≤ 28 000 kg
- Труба+Профиль ≤ 11 000 kg
- Лист ≤ 18 000 kg
- Max 4 konteyner/kun, 20 konteyner/oy
- Ustunlik: 🔴КРИТИК > 🟡ПАСТ > 🟢НОРМА → urgentlik (kun)
- Tovar nomi `YANGI:` prefiksi bilan qaytsa — inventarda yo'q yangi tovar

### 2.5 Xitoy fayl formatlari

| Format | Signal ustunlar | Qayta ishlash |
|--------|----------------|---------------|
| Труба/Профиль | 规格, 长度, 库存 (Ombor) | `_parse_truba_profil_xitoy()` |
| Лист | 品号, 颜色, 数量 | `_parse_list_xitoy()` |
| Oddiy | Товар, Миқдор | To'g'ridan ustun |

---

## 3. FAYL STRUKTURASI (bot_holat/)

| Fayl | Maqsad |
|------|--------|
| `buyurtma_{kanal}.json` | Tasdiqlangan buyurtmalar (Excel olishda ayiriladi) |
| `draft_{kanal}.json` | Excel yaratilganda tovar ro'yxati (buyurtma_tekshir lookup) |
| `xitoy_{kanal}.json` | Xitoy ostatka: `tovarlar` (zakaz K) + `ombor` (tayyor L) |
| `pending_{kanal}_{user_id}.json` | Preview ko'rsatilgan lekin tasdiqlanmagan buyurtma |

---

## 4. EKRANLAR VA BACK_MAP

```
main
├─ order → order_channel
├─ load  → load_channel
├─ status → (karta / yolda)
├─ search
└─ settings
```

`BACK_MAP`:
- `order_channel` → `order`
- `load_channel` → `load`
- `order` → `main`
- `load` → `main`

---

## 5. MUHIM TEXNIK JIHATLAR

### xitoy_ostatka_oqi() — ombor vs tovarlar farqi
- `known_map` = K ustun (zakaz miqdori, `tovarlar` kalitiga saqlanadi)
- `ombor_map` = L ustun (tayyor/ready, `ombor` kalitiga saqlanadi)
- **Buyurtma Excel** → `xitoy_ostatka` = `tovarlar` (zakaz K, buyurtmani kamaytirish uchun)
- **Yuklatish rejasi** → `akkum` = `ombor` (tayyor L, nimani yuklash mumkin)

### grafik_qidirish() — format
- Truba: `51 -> 0,9 -> 5,8 -> 201` (diametr → stenka → uzunlik → marka)
- Profil: `20х20 -> 0,7 -> 6 -> 201` (o'lcham → stenka → uzunlik → marka)
- List: `0,8 -> 1220 -> 201` (qalinlik → kenglik → marka)
- Balasina: tovar nomini to'liq yozing

### buyurtma_tekshir() — case-insensitive (tuzatilgan)
- "Tovar nomi" sarlavhasi `.lower()` bilan taqqoslanadi
- Varaq nomi (Труба/Профиль/Лист...) katta/kichik harf ahamiyatsiz
