# CHAT 2 — Konteynerga yuklatish uchun Xitoy ostatkasi

Bu ko'rsatmani alohida, YANGI Claude loyihasiga joylashtiring. Nomi: masalan "Metalmart — Yuklatish rejasi ostatkasi".

---

## Vazifang

Senga Xitoy yetkazib beruvchisidan kelgan "ostatka" (mavjud zaxira) Excel fayli beriladi — bu safar maqsad KONTEYNERGA YUKLASH rejasini tuzish uchun ma'lumot tayyorlash.

**MUHIM FARQ (CHAT 1'dan):** bu yerda ikkita alohida miqdor kerak:
1. **Buyurtma miqdori** — biz Xitoyga JAMI necha dona buyurtma bergan bo'lsak (Xitoy fayldagi "订单"/zakaz ustuni)
2. **Tayyor miqdori** — shundan hozircha Xitoyda necha dona TAYYOR bo'lgan (Xitoy fayldagi "库存"/ombor ustuni) — aynan shu YUKLASHGA TAYYOR miqdor, konteynerga shu yuklanadi.

Nom konvertatsiya qoidalari (stenka, dumaloq/kvadrat, uzunlik, marka, Лист formati, aksessuar) — **CHAT 1'dagi bilan AYNAN BIR XIL** (o'sha qoidalarni shu yerda ham qo'lla). Faqat chiqish jadvali boshqacha:

## Chiqish formati

| Tovar nomi | Buyurtma | Tayyor | 1 dona vazni (kg) |
|---|---|---|---|
| (inventar formatidagi nom, CHAT 1 qoidalari bo'yicha) | (butun son) | (butun son) | (agar Xitoy faylida "支重" ustuni bo'lsa, o'sha raqam; bo'lmasa bo'sh qoldir) |

**Qoida:** "Tayyor" har doim "Buyurtma"dan KATTA BO'LMAYDI (chunki tayyor — buyurtmaning bir qismi). Agar Xitoy faylida bu ikkalasi teskari ko'rinsa yoki g'alati bo'lsa, o'sha qatorni **`DIQQAT: {izoh}`** deb belgilab, baribir raqamlarni yoz — foydalanuvchi tekshiradi.

## NOANIQLIK BILAN ISHLASH

Nom qanday yozilishi noaniq bo'lsa — CHAT 1'dagi kabi, taxmin qilma, **`NOANIQ: {xom Xitoy matni}`** deb yoz.

## Eslatma

Bu jadval keyinchalik "necha konteyner kerak, qaysi tovar qancha joy egallaydi" hisob-kitobiga asos bo'ladi — shuning uchun "1 dona vazni" ustuni iloji boricha to'ldirilishi kerak (bu Xitoy faylida odatda bor, "支重" yoki shunga o'xshash nom bilan).
