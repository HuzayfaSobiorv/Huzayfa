# NEJAVIYKA — Buyurtma berish jarayoni (logika hujjati)

> **Maqsad:** Bot orqali Xitoyga yangi zakaz yuborishdan oldin qancha buyurtma kerakligini aniq hisoblash.

---

## 1. Bosqich: Xitoy Ostatka faylini yuklash

Foydalanuvchi **"📄 Buyurtma Excel olish"** tugmasini bosganda bot darhol Excel yubormaydi.

**Agar avvalgi Xitoy ostatka mavjud bo'lsa:**
> Bot so'raydi: "Mavjudni ishlatsin" yoki "Yangi yuklash"

**Agar Xitoy ostatka yo'q bo'lsa:**
> Bot so'raydi: "Ha, bor" yoki "Yo'q — shundayicha ber"

Foydalanuvchi bir yoki bir necha Xitoy ostatka Excel faylini yuborishi mumkin.
Har fayldan keyin bot so'raydi: "➕ Yana fayl" yoki "✅ Tayyor — Excel yaratish"

---

## 2. Xitoy Ostatka fayli nimani ifodalaydi

Xitoy bizga Excel yuboradi — unda:

| Ustun      | Ma'no                                                            |
|------------|------------------------------------------------------------------|
| Mahsulot   | Tovar nomi (pipe/profil/list formati)                            |
| Zakaz (K)  | Bizning eski zakazimiz — Xitoy bilgan va ishlab chiqarayotgan    |
| Tayyor (L) | O'sha zakazdan hozir jo'natishga tayyor bo'lgani                 |

Bot bu faylni `xitoy_{kanal}.json` ga saqlaydi:
- `tovarlar` → K ustun (buyurtma hisoblash uchun ayiriladi)
- `ombor`    → L ustun (yuklatish rejasi uchun)

---

## 3. Buyurtma hisoblash — to'liq formula

```
buyurtma_kerak = max(0, sotuv55 + min_zaxira − qoldiq − yolda − xitoy_K − tasdiqlangan)
```

| Komponent       | Manba                                  | Ta'rif                                              |
|-----------------|----------------------------------------|-----------------------------------------------------|
| `sotuv55`       | Power BI trenddan                      | 55 kunlik kutilayotgan sotuv (min_zaxira/30 * 55)   |
| `min_zaxira`    | Min_Zaxira.xlsx                        | Belgilangan minimal zaxira miqdori                  |
| `qoldiq`        | Tarix/*.xlsx (bugungi)                 | Omborlardagi joriy miqdor                           |
| `yolda`         | konteynerlar/*.xlsx                    | Yo'ldagi konteynerlarda kelayotgan miqdor           |
| `xitoy_K`       | Xitoy ostatka fayl, K ustun            | Xitoy BILGAN va ishlab chiqarayotgan zakaz           |
| `tasdiqlangan`  | bot_holat/buyurtma_{kanal}.json        | Biz yangidan tasdiqladik, Xitoy hali bilmaydi       |

---

## 4. xitoy_K va tasdiqlangan — farqi va munosabati

**Muhim tushuncha:**

```
xitoy_K       = Xitoy AVVALGI zakazimizni biladi va qayta ishlayapti
tasdiqlangan  = Biz YANGI zakaz berdik lekin Xitoy hali ostatkaga KO'SHMAGAN
```

Ular **bir-birini qoplamaydigan, to'ldiruvchi** komponentlar:
- Ikkisi ham `buyurtma_kerak` dan ayiriladi
- Natijada ikki tomondan ham qoplanilgan miqdor hisobdan chiqariladi

**Yangi ostatka kelganda nima bo'ladi:**
```
Xitoy yangi ostatka yubordi
    → xitoy_K endi bizning tasdiqlangan zakazni ham o'z ichiga oladi
    → buyurtma_tozala(kanal) → tasdiqlangan.json o'chadi
    → Chunki endi tasdiqlangan = xitoy_K ning ichida
```

---

## 5. "Mavjudni ishlatsin" va "Yangi yuklash" farqi

| Tanlov              | Nima qiladi                                               | tasdiqlangan |
|---------------------|-----------------------------------------------------------|--------------|
| **Mavjudni ishlatsin** | Diskdagi xitoy JSON dan o'qiydi, buyurtma hisoblaydi   | ✅ Hisobga oladi |
| **Yangi yuklash**   | Eski xitoy JSON o'chadi, yangi fayl kutiladi              | ✅ Hisobga oladi |
| **Ha, bor** (yo'q bo'lganda) | Yangi fayl yuklaydi                           | ✅ Hisobga oladi |
| **Yo'q** (yo'q bo'lganda)    | xitoy_K = 0 deb hisoblanadi                  | ✅ Hisobga oladi |

> **Eslatma:** Tasdiqlangan buyurtma (buyurtma_{kanal}.json) BARCHA holatlarda
> hisobga olinadi. U faqat yangi xitoy ostatka yuklanganida o'chadi.

---

## 6. Buyurtma Excel foydalanuvchiga yuboriladi

Bot 6 ta varaqli Excel yuboradi:
- Труба, Профиль, Лист, Баласина, Стойка, Аксессуар
- Har varaqda: Tovar nomi | Uzunlik | Buyurtma miqdori

Foydalanuvchi Excelni olib o'zgartirishi mumkin.

---

## 7. Tasdiqlash — "Yuborish"

> **QOIDA:** "Yuklash" so'zi ishlatilmaydi. Faqat **"Yuborish"** ishlatiladi.

Foydalanuvchi:
1. O'zgartirgan Excelni botga **yuboradi**
2. Bot preview ko'rsatadi + "✅ Tasdiqlayman" / "❌ Bekor"
3. "Tasdiqlayman" → `buyurtma_{kanal}.json` ga saqlanadi

Bu fayl = "Men Xitoyga BU tovarlarni buyurtma berdim, ular hali bilishmaydi"

---

## 8. Tasdiqlangan fayl hayot davri

```
Yaratiladi:   Foydalanuvchi "Tasdiqlayman" ni bosganda
O'chadi:      FAQAT yangi Xitoy Ostatka fayli yuklanganida

Sababi:       Yangi ostatka = Xitoy bizning zakazimizni qabul qildi
              va K ustuniga qo'shdi → ikki marta ayirilmasin
```

---

## 9. Holat diagrammasi

```
[Buyurtma Excel olish]
        ↓
  Xitoy JSON bormi?
  ├─ Ha  → [Mavjudni ishlatsin] yoki [Yangi yuklash]
  └─ Yo'q → [Ha, bor → fayl yuklash] yoki [Yo'q → xitoy_K=0]
        ↓
[Excel generatsiya]
  buyurtma = sotuv55 + min − qoldiq − yolda − xitoy_K − tasdiqlangan
        ↓
[Excel foydalanuvchiga yuborildi]
        ↓
[Foydalanuvchi o'zgartiradi → yuboradi]
        ↓
[Preview + Tasdiqlayman / Bekor]
        ↓ (Tasdiqlayman)
[buyurtma_{kanal}.json saqlandi]
        ↓ (keyingi ostatka yuklanganida)
[buyurtma.json o'chadi — xitoy_K da aks etgan]
```

---

## 10. Xato bo'lish holatlari

| Holat                              | Natija                                          |
|------------------------------------|-------------------------------------------------|
| Ikki marta tasdiqlash              | Miqdorlar qo'shiladi ⚠️ ogohlantirish chiqadi  |
| Xitoy fayli o'qilmasa              | Xato xabar + holat saqlanadi                   |
| Bot qayta ishga tushsa             | pending_zakaz diskdan tiklanadi                 |
| buyurtma_kerak = 0                 | "Hozircha buyurtma kerak emas" xabari           |
