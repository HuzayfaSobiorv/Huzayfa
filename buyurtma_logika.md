# NEJAVIYKA — Buyurtma berish logikasi (2026-07-09 holati)

> Eski versiya (statik `sotuv55` formula) butunlay bekor qilindi.
> Hozirgi hisob-kitob **zanjir-simulyatsiya** (`kamomat_engine.zanjir_sim`) asosida.

---

## 1. Umumiy oqim (o'zgarmagan qism)

1. **"📄 Буюртма Excel олиш"** bosiladi.
2. Bot Xitoy ostatka holatini so'raydi: mavjud bo'lsa "Mavjudni ishlatsin" / "Yangi yuklash", bo'lmasa "Ha, bor" / "Yo'q".
3. Qaror asosida `draft_buyurtma_yubor()` → `services.asosiy_styled_excel_yarat()` chaqiriladi.
4. Excel foydalanuvchiga yuboriladi, `draft_{kanal}.json` ga tovar ro'yxati saqlanadi (keyingi tekshirish uchun).
5. Foydalanuvchi Excelni o'zgartirib qaytarib yuboradi → preview → **Tasdiqlayman** → `buyurtma_{kanal}.json` ga yoziladi (eskisi bilan qo'shilib — yig'iladi).

---

## 2. Buyurtma miqdorini hisoblash — YANGI mantiq

### 2.1 Kunlik sotuv — yagona qoida

```
kunlik_sotuv = min_zaxira / KUNLIK_SOTUV_BOLISH      (common.py, = 30)
```

Bu son **hamma joyda** bir xil: buyurtma hisobi, kamomat holati, qidiruv/grafik.
`KELISH_KUNI` (=55, haqiqiy jismoniy yetib kelish vaqti) bilan **adashtirilmaydi** — u alohida konstanta, faqat KRITIK/PAST klassifikatsiya va grafik x-o'qi uchun ishlatiladi.

### 2.2 Zanjir-simulyatsiya (`kamomat_engine.zanjir_sim`)

Har bir tovar uchun **real konteyner kelish sanalari** bilan kun-ma-kun simulyatsiya qilinadi (statik "jami yo'lda" emas):

- Boshlang'ich: bugungi `qoldiq`.
- Har kuni `kunlik_sotuv` miqdorda kamayadi.
- Har konteyner o'z `Кун_Қолди` kunida `Миқдор`ni qo'shadi (manba: Power BI "Контейнерлар" varag'i, `Холат != "КЕЛДИ ✅"`).
- Simulyatsiya davomida eng past nuqta (`min_nuqta`) va `min_zaxira`dan pastga tushgan kun (`uzilish_kun`) qayd etiladi.

Natija:
```
taklif_A = min_zaxira - (qoldiq + yo'ldagi_jami)        # min darajaga qaytarish
taklif_B = min_zaxira - min_nuqta   (agar min_nuqta < min_zaxira)   # zanjirdagi bo'shliqni yopish
taklif   = max(taklif_A, taklif_B), 50 ga yaxlitlangan
xavf     = KRITIK / PAST / NORMA / MEYOR_YOQ
```

**Nega kerak edi:** eski statik formula konteynerlar orasidagi bo'shliqni (masalan 23–42 kun oralig'ida hech narsa kelmasligini) hisobga olmasdi. Endi shu bo'shliq davridagi eng past nuqta ham tekshiriladi — real tanqislik yo'qolib qolmaydi.

Kod: `Generate_Asosiy_order.py::load_data()` — "Контейнерлар" varag'idan `kont_map: {tovar: [(kun_qoldi, miqdor), ...]}` quradi → `calculate()` har qatorga `zanjir_sim()` chaqiradi.

### 2.3 Keyingi ayirishlar (`services.py::asosiy_styled_excel_yarat`)

`zanjir_sim` natijasidan ketma-ket ayiriladi:

1. **Xitoy ostatka K ustuni** — Xitoy allaqachon bilib ishlab chiqarayotgan zakaz.
2. **Tasdiqlangan buyurtma** (`buyurtma_{kanal}.json`) — biz yangi tasdiqladik, Xitoy hali K ustuniga qo'shmagan.

```
yakuniy_buyurtma = max(0, zanjir_taklif - xitoy_K - tasdiqlangan)
```

Ikkisi ham bir-birini to'ldiruvchi, bir-birini qoplamaydi — ikkisi ham ayiriladi.

---

## 3. Xitoy ostatka va tasdiqlangan — hayot davri

- `xitoy_{kanal}.json` — diskda **doimiy** qoladi, avtomatik o'chmaydi.
- `buyurtma_{kanal}.json` (tasdiqlangan) — diskda **doimiy** qoladi, avtomatik o'chmaydi.
- Ikkisi ham o'chadigan **YAGONA** holat: foydalanuvchi keyingi safar **"Yangi yuklash"**ni tanlaganda (fayl kelishidan OLDIN, tugma bosilgan zahoti) — `handlers.py` da eski xitoy JSON unlink qilinadi va `buyurtma_tozala(kanal)` chaqiriladi.
- **Sabab:** "Tasdiqlayman" bosish = zakaz haqiqatda Xitoyga jo'natilgan degani. Demak keyingi Xitoy ostatka fayli kelganda, u yangi K ustunida bizning avvalgi tasdiqlangan zakazimizni ham allaqachon aks ettiradi — shu sabab eski tasdiqlangan endi ortiqcha, ikki marta ayirilmasligi uchun tozalanadi.
- "Mavjudni ishlatsin", "Hisobsiz ber", "Yo'q" — bularning hech birida hech narsa o'chmaydi.

---

## 4. Hozirgi qamrov chegarasi

Faqat **Труба / Профиль / Лист** uchun ishlaydi va tekshirilgan (13900 dona Ф-51 misolida tasdiqlangan).
Aksessuar toifalari (Баласина, Стойка, Чашка, Соққа, Шар, Отвод) hali eski/oddiy hisobda — ularga hozircha tegilmayapti (foydalanuvchi talabi bilan).

---

## 5. Keyinga qoldirilgan (hali qo'shilmagan)

- ABC toifaga qarab himoya buferi: A=+15%, B=+10%, C=+5% (mexanizm buffersiz holda tayyor, keyin ustiga qo'shiladi).
- Konteyner **qachon yuklanishi kerak** — vaqt jadvali/rejasi (alohida masala, zanjir-simulyatsiyadan farqli).
