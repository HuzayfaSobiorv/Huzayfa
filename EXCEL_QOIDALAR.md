# NEJAVIYKA — Buyurtma Excel Dizayn Qoidalari

Sana: 2026-06-26  
Fayl: `Generate_Asosiy_order.py`

---

## 1. Stranitsalar tartibi (CATEGORIES)

Excel faylida quyidagi tartibda varaqlar bo'ladi:

| # | Varaq nomi       | Izoh |
|---|------------------|------|
| 1 | Труба            | Payvandli truba (Ф-диаметр ст stenka) |
| 2 | Профиль          | Kvadrat/to'g'riburchak profil (Пр. AxB ст stenka) |
| 3 | Лист             | Varaq (Лист-qalinlik format) |
| 4 | Балясина         | Balyasina (dekorativ ustunyoqalar) |
| 5 | Безшовный труба  | Bezshvovniy (tikishsiz) truba — Балясинадан KEYIN turadi |
| 6 | Стойка           | Stojka |
| 7 | Шар              | Shar |
| 8 | Соқка            | Sokka |
| 9 | Чашка            | Chashka — **Қузиқорин ham shu varaqqa kiritiladi** |
|10 | Полировка        | Polirovka aksessuar (Намат, Мелкий, Грубый, Капрон, Шлифовка) |

> **Қузиқорин alohida varaq emas** — Чашка varag'ining ichida ko'rsatiladi.
> Чашка varaqida avval Чашка mahsulotlari, keyin `── Қузиқорин ──` ajratuvchisi, keyin Қузиқорин mahsulotlari.

---

## 2. Har bir varaqdagi ustunlar

| Ustun | Nomi      | Kenglik | Format |
|-------|-----------|---------|--------|
| A     | Tovar nomi | 50      | 15pt Calibri bold qora |
| B     | Uzunlik   | 12      | 15pt Calibri bold qizil, markazlashgan |
| C     | Buyurtma  | 14      | 15pt Calibri bold qizil, markazlashgan |

> Uzunlik (B ustun) faqat **Труба va Профиль** uchun to'ldiriladi (m da: 5.8, 6.0 ...).  
> Qolgan kategoriyalar uchun B bo'sh qoladi (yoki kelishib olinadi).

---

## 3. Qatorlar tuzilishi (Труба, Профиль, Лист uchun — sorted_mode)

```
[1] Kategoriya banner   → to'q ko'k fon, oq matn, 30pt balandlik
[2] Ustunlar sarlavhasi → ko'k fon, oq matn, 26pt
[3] Surface ajratuvchi  → ── Матовый ── (agar mavjud bo'lsa)
[4] Tovar qatori        → alternating och ko'k/oq, 26pt
...
```

**Surface tartibi:** `"" (Oddiy)` → `Матовый` → `Глянцевый` → `Чёрный` → `Голд` → `Цветной`  
**Marka tartibi:** `201` → `304` → `430` → `316` → `321` → `""` (noma'lum)

---

## 4. Format ajratuvchi (FAQAT Лист uchun)

```
── КАТТА ФОРМАТ (>1250 мм) ──   ← to'q ko'k (#1A3A5C)
── КИЧИК ФОРМАТ (≤1250 мм) ──   ← biroz ochroq ko'k (#2C4A6E)
```

> Bu ajratuvchi **FAQAT Лист varaq**iga tegishli.  
> Труба va Профиль varaqlari uchun bunday ajratuvchi **YO'Q**.

---

## 5. Tovar nomidagi XITOY_DELTA

Excelda ko'rsatiladigan tovar nomi **Xitoy buyurtma nomi** bo'ladi:  
- **FAQAT Лист** uchun: qalinlik `0.05 mm` kamaytirib ko'rsatiladi  
  → `Лист-0,8 (1220х2440)` → Excel da `Лист-0,75 (1220х2440)`  
- **Труба va Профиль** uchun: o'zgarishsiz — inventardagi nom bilan bir xil

---

## 6. Чашка + Қузиқорин birlashtirilishi

Чашка varaqida ikki bo'lim bo'ladi:

```
[Banner] ЧАШКА — дд.мм.йй
[Sarlavha] Tovar nomi | Uzunlik | Buyurtma
  [Чашка mahsulotlari — alphabetik]
── Қузиқорин ──          ← ajratuvchi qator (SURFACE_BG rang)
  [Қузиқорин mahsulotlari — alphabetik]
```

---

## 7. Bezshvovniy truba (Безшовный труба)

- Inventardagi tovar nomi: `Безш.` yoki `Безшовный` prefiks bilan
- Kategoriya: `Безшовный труба`
- Varaqda tartib: sorted_mode=False (oddiy alphabetik), Балясинадан keyin
- Uzunlik ustuni: to'ldiriladi (5.8, 6.0 m ...)

---

## 8. Шар

- Inventardagi tovar nomi: `Шар` so'zi bilan boshlanadi  
  (Шаркона, Шарнир, Шаршара **kirmaydi**)
- sorted_mode=False (alphabetik)
- Стойкадан keyin, Соккадан oldin

---

## 9. Rang palitrasi

| Element              | Rang (HEX)  |
|----------------------|-------------|
| Kategoriya banner    | `1F4E79`    |
| Ustun sarlavha       | `2E75B6`    |
| Surface ajratuvchi   | `D6E4F0`    |
| Qator (toq)          | `DAEEF3`    |
| Qator (juft)         | `F5FBFC`    |
| Katta format ajrat.  | `1A3A5C`    |
| Kichik format ajrat. | `2C4A6E`    |
| Tovar nomi matn      | `000000`    |
| Uzunlik/Buyurtma     | `C00000`    |
