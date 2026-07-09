# NEJAVIYKA Bot — Tugmalar, oqimlar va holat xaritasi (to'liq ma'lumotnoma)

Bu hujjat botning HAR bir tugmasi, menyusi va oqimini, ularning hozirgi holatini, ma'lumot manbasini va kod joyini (fayl:funksiya) ko'rsatadi. Asos: `keyboards.py`, `texts.py`, `ui.py`, `config.py`, `Bot.py` va `handlers.py` (to'g'ri versiyasi `git show HEAD:handlers.py` dan olindi — ish papkasidagi nusxa hozircha kesilgan/buzuq holatda).

## Umumiy arxitektura (kalit tushunchalar)

- **Kirish nuqtasi:** `Bot.py` → `start`, `callback_handler` (barcha inline tugmalar), `text_keldi` (barcha reply-tugma va matn), `fayl_keldi` (.xlsx fayllar), va admin komandalar. Hamma handler `filters.ChatType.PRIVATE` bilan cheklangan — faqat shaxsiy chatda. Yagona istisno `/chatid` (guruhda ham ishlaydi).
- **Ekran/holat mashinasi:** `context.user_data["screen"]` joriy ekranni, `context.user_data["kutilmoqda"]` esa "keyingi qadamda nima kutilyapti" (fayl, matn, sana va h.k.) ni saqlaydi. Ekranlar `ui.py:build_screen`/`go_screen` orqali chiziladi, tugma→amal moslamasi `ui.py:get_action` da (`MAP` lug'ati). "Orqaga" navigatsiyasi `config.py:BACK_MAP` orqali.
- **Ma'lumot manbalari:**
  - `chiqish/NEJAVIYKA_POWER_BI.xlsx` — asosiy hisobot fayli (varaqlar: `Инвентар`, `Контейнерлар`, `Критик`, `Йўлдаги_Контейнерлар` va b.). `config.py`da `DATA_FILE`. RAM keshda (`config.py:xlsx_refresh`, 5 daqiqa TTL). `main.py` ishlab chiqaradi.
  - `main.py` kirishlari: `Tarix/` (kunlik qoldiq), `Minimal_zaxiralar/Min_Zaxira.xlsx` (minimal zaxira), `konteynerlar/xitoy_parsed/` (yo'ldagi/kelgan konteyner fayllari). Chiqishi: `chiqish/NEJAVIYKA_POWER_BI.xlsx` (+ nusxa ildizga).
  - `bot_holat/` JSON holatlari: `buyurtma_{kanal}.json` (tasdiqlangan buyurtma), `xitoy_{kanal}.json` (Xitoy ostatka), `pending_*` (tasdiq kutayotgan), `qoshilgan_konteynerlar.json` (konteyner ISO tarixi), `whitelist.json` (ruxsat), `xitoy_tuzatishlar.json` (nom mapping).
- **Kanallar:** `asosiy`, `sex`, `osh`.

  **JIDDIY XATO (tasdiqlangan):** `config.py:114` —
  ```
  CH_KEY = {"asosiy": "ch_asosiy", "sex": "ch_sex", " osh": "ch_osh"}
  ```
  "osh" kalitida **oldida bo'sh joy bor** (`" osh"`), lekin `ui.py:192,201` da kanal `"osh"` (bo'sh joysiz) sifatida saqlanadi va `handlers.py:1030,1080,1092`, `ui.py:131,136,353,383`, `services.py:691` da `CH_KEY[kanal]` shaklida ishlatiladi. Natijada `CH_KEY["osh"]` → **KeyError**. Bu O'sh (Qirg'iziston) kanali orqali "Buyurtma yig'ish" yoki "Konteyner yuklash" tanlanganda botni qulatadi. **Asosiy va Tsex kanallari ishlaydi, O'sh — buzuq.** Tuzatish bir qatorlik: `" osh"` → `"osh"`.

---

## Kirish va ro'yxatdan o'tish (whitelist)

**Holati: to'liq ishlayapti.**

- `/start` (`handlers.py:start`): `kirish_ruxsati(uid)` (`services.py`, `whitelist.json` + `ADMIN_IDS`) tekshiriladi. Ruxsat yo'q bo'lsa — foydalanuvchiga "kirish huquqi yo'q" xabari, SUPER_ADMIN_ID ga "Yangi kirish so'rovi" + `/adduser <id>` yuboriladi. Ruxsat bor va til tanlangan bo'lsa — `main` ekran; til tanlanmagan bo'lsa — til tanlash inline (`til_ikb`).
- Til inline (`callback_data`: `lang:cyr`, `lang:lat`) — tilni saqlaydi va `main` ekranga o'tadi.

**Admin/foydalanuvchi farqi:** `ui.py:_is_admin` — `ADMIN_IDS` (`.env`dan). Admin bo'lsa to'liq menyu (`main_kb`), oddiy filial foydalanuvchisi bo'lsa soddalashtirilgan menyu (`main_kb_user`).

---

## Admin komandalar (faqat matn buyruq, tugma emas)

Barchasi `handlers.py`, faqat `SUPER_ADMIN_ID` uchun (chatid — `ADMIN_IDS`):

| Komanda | Holati | Vazifa | Kod |
|---|---|---|---|
| `/adduser <id>` | ishlaydi | whitelist ga qo'shadi, foydalanuvchiga xabar yuboradi | `adduser_cmd` → `whitelist_qosh` |
| `/removeuser <id>` | ishlaydi | whitelist dan o'chiradi | `removeuser_cmd` → `whitelist_ochir` |
| `/users` | ishlaydi | Adminlar + whitelist ro'yxatini ko'rsatadi | `users_cmd` → `whitelist_yuklash` |
| `/chatid` | ishlaydi | Joriy chat/topic ID sini beradi (.env sozlash uchun; guruhda ishlaydi) | `chatid_cmd` |

---

## Bosh menyu (Asosiy menyu)

### Admin varianti — `keyboards.py:main_kb`
- **📥 Buyurtma yig'ish** (`b_order`) → `order` ekran. Ishlaydi.
- **🚛 Konteyner yuklash** (`b_load`) → `load` ekran. Ishlaydi.
- **♻️ Konteynerlar** (`b_konteyner`) → `konteyner` ekran. Ishlaydi.
- **🔍 Qidiruv** (`b_karta`) → `karta` amali. Ishlaydi.
- **⚙️ Sozlamalar** (`b_settings`) → `settings` ekran. Ishlaydi.

### Foydalanuvchi (filial) varianti — `keyboards.py:main_kb_user`
- **🚛 Yo'ldagi yuklar** (`b_yolda_excel`) → `yolda_excel` amali. Ishlaydi.
- **🔍 Qidiruv** (`b_karta`). Ishlaydi.
- **⚙️ Sozlamalar** (`b_settings`). Ishlaydi.

**Ishlatilmayotgan/eskirgan tugmalar (o'lik kod):**
- `b_status` ("📈 Holat"), `b_search` ("🔍 Qidiruv" alohida), `b_yolda` ("🚢 Yo'ldagi holat") — hech bir reply-klaviaturada ko'rsatilmaydi. `status_kb` ekran umuman ochilmaydi.
- `text_keldi`dagi `action == "yolda"` → **`pass`** (bo'sh, o'lik).
- `b_upload` ("📎 Xitoy Excelini yuklash") — hech qaysi klaviaturada yo'q; `action == "upload"` shoxi mavjud, lekin unga hech qaysi tugma yo'naltirmaydi → o'lik kod.

---

## Buyurtma yig'ish oqimi (Buyurtma Excel + tasdiqlash)

**Umumiy holati: to'liq ishlayapti (Asosiy/Sex uchun); O'sh — CH_KEY xatosi tufayli buzuq.**

Ekranlar: `main` → `order` (kanal tanlash: `order_kb`) → `order_channel` (`order_channel_kb`).

### 1) Kanal tanlash — `order` ekran
- **🏢 Asosiy** / **🏭 Tsex** / **🇰🇬 O'sh** → `order_channel` ekran, `kanal` saqlanadi. Asosiy/Tsex ishlaydi; **O'sh → KeyError.**

### 2) Amal tanlash — `order_channel` ekran
- **📄 Buyurtma Excel olish** (`b_excel`) → amal `excel`.
  - Avval `xitoy_yuklash(kanal)` bilan Xitoy ostatka JSON borligini tekshiradi: JSON bor va tovarlari bo'lsa → inline `xitoy_mavjud_ikb`; bo'sh yoki yo'q bo'lsa → `xitoy_sorash_ikb`.
  - Yakuniy natija: `ui.py:draft_buyurtma_yubor` → `services.py:draft_excel_yarat` → `asosiy_styled_excel_yarat` → `Generate_Asosiy_order`. Foydalanuvchiga **`Buyurtma_taklif_{kanal}.xlsx`** yuboriladi. Ma'lumot: `NEJAVIYKA_POWER_BI.xlsx` (kamomat) + `xitoy_{kanal}.json` (ostatka) + `buyurtma_{kanal}.json` (allaqachon berilgan buyurtma chiqarib tashlanadi).
  - Holati: **ishlaydi.**
- **✅ Tasdiqlangan buyurtmani yuborish** (`b_tasdiq`) → amal `tasdiq`.
  - Foydalanuvchidan to'ldirilgan Excel kutiladi → `services.py:buyurtma_tekshir` (format tekshiruvi, nom lookup) → xato bo'lsa `tasdiq_err`; to'g'ri bo'lsa `zakaz_preview_text` + inline `zakaz_tasdiq_ikb`.
  - Holati: **ishlaydi.**

### Xitoy ostatka inline savoli — `callback_handler` `xitoy:*`
- `xitoy:ha:{kanal}` — 2 fayllik oqim: Труба/Профиль, keyin Лист. `parsers.py:xitoy_ostatka_oqi`/AI-format bilan parse, `bot_holat/xitoy_{kanal}.json` ga yoziladi, so'ng draft Excel.
- `xitoy:yoq`/`xitoy:hisob_siz` — hisobga olinmaydi, darhol draft Excel.
- `xitoy:ishlatsin` — mavjud JSON ishlatiladi.
- `xitoy:yangi` — eski JSON + tasdiqlangan buyurtma o'chadi, yangi fayl so'raladi.
- `xitoy:yana_f`/`xitoy:tayyor` — ko'p fayl yig'ish.
- Holati: **ishlaydi.**

### Zakaz tasdiqlash 2-bosqich
- `zakaz_ok:{kanal}` → `buyurtma_saqlash` (`bot_holat/buyurtma_{kanal}.json`). Ishlaydi.
- `zakaz_no` → pending tozalanadi. Kichik nuqson: fallback ro'yxati `["asosiy","cex","osh"]` — "cex" emas "sex" bo'lishi kerak edi (ta'siri yo'q, chunki mavjud bo'lmagan fayl e'tiborsiz qoldiriladi).

---

## Konteyner yuklash oqimi (Yuklatish rejasi)

**Umumiy holati: to'liq ishlayapti (Asosiy/Sex); O'sh — CH_KEY xatosi.**

### 1) Kanal tanlash — `load` ekran
- **🏢 Asosiy** / **🏭 Tsex** / **🇰🇬 O'sh** → `load_channel` ekran, akkumulyatorlar (`xitoy_akkum`, `ombor_akkum`, `vazn_akkum`) reset qilinadi. Asosiy/Sex ishlaydi; **O'sh → KeyError.**

### 2) Fayl yuborish va hisoblash — `load_channel` ekran
- Xitoy tayyor ro'yxati (.xlsx) yuboriladi, bir nechta fayl akkumulyatorga qo'shiladi. Har fayldan keyin inline `xitoy_yana_ikb` ("➕ Yana fayl"/"▶️ Hisoblashni boshlash").
- **▶️ Hisoblash**: `yuklatish_rejasi.main_with_data(kanal, ombor_akkum, xitoy_vazn=vazn_akkum)` (`vazn_hisobla`, `Yuklama_optimal`, `yuklatish_rejasi`). Kamomat manbasi: `NEJAVIYKA_POWER_BI.xlsx`; yuk tarkibi: Xitoy ombor akkumulyatori.
  - Natija: **`Yuklatish_rejasi.xlsx`** + statistika. Maxsus holatlar (`KERAK_YOQ`, `OMBOR_BOʻSH`, `MOS_YOQ|...`, `STATS:...`) to'g'ri ishlanadi.
  - Holati: **ishlaydi** (2026-07-08 da bo'sh edi — keyinchalik tuzatilgan, hozir to'liq).

---

## Qidiruv / Tovar kartasi oqimi

**Holati: to'liq ishlayapti.**

- **🔍 Qidiruv** (`b_karta`) → `search` ekran + inline kategoriya klaviaturasi (`grafik_kat_ikb`).
- Kategoriya inline (`karta_kat:*`): Труба, Профиль, Лист, Баласина, Стойка, Чашка, Қузиқорин, Шар, Соққа, Ойна держатель — har biri mos format so'rovini so'raydi.
- **🔍 Umumiy qidiruv** (`karta_umumiy`) — kategoriyasiz.
- Natija: `grafik_qidirish` → bitta bo'lsa darhol karta, ko'p bo'lsa inline tanlash (`karta_tovar:{i}`).
- Karta: `ui.py:grafik_ko_rsatish` — matn (qoldiq, min, yo'lda jami, yo'ldagi konteynerlar, uzilish xavfi) + Труба/Профиль/Лист uchun grafik rasm (`kamomat_engine:grafik_chiz`). Manba: `NEJAVIYKA_POWER_BI.xlsx` (RAM kesh).

**Eslatma:** `texts.py:search_stub`/`kritik_stub` va `ui.py:kamomat_ko_rish` hech qayerda chaqirilmaydi — eski/o'lik kod. Kamomat hisobi endi bevosita draft Excel (`Generate_Asosiy_order`) ichida bajariladi.

---

## Konteynerlar bilan ishlash (♻️ Konteynerlar ekran)

### A) 🚛 Yo'ldagi yuklar (`yolda_excel`)
**Holati: to'liq ishlayapti.** `ui.py:yolda_ko_rish` → `yolda_excel.py:yolda_excel(DATA_FILE)`. Manba: `NEJAVIYKA_POWER_BI.xlsx`. Natija: **`Yolda.xlsx`** (kelish sanasi bo'yicha tartiblangan). Filial foydalanuvchisi ham ko'radi.

### B) 🚢 Yo'lga konteyner qo'shish (`yolga_kont`, faqat admin)
**Holati: to'liq ishlayapti.** Ko'p bosqichli oqim:
1. Труба/Профиль `装箱单` fayli so'raladi (aksessuar fayl bo'lsa 2-bosqich o'tkazib yuboriladi).
2. Лист `出货清单` fayli. Ikkalasi `konteyner_qosh.xitoy_yuklar_oqi` bilan solishtiriladi.
3. Sana/ISO filtrlari — takror qo'shilmasligi `qoshilgan_konteynerlar.json` orqali ta'minlanadi.
4. Draft xulosa + tahrirlanadigan **`Yangi_konteynerlar.xlsx`**, notanish tovar ogohlantirishi.
5. Tasdiqlash: `kont:ha` → `konteyner_qosh.konteyner_xlsx_yarat` (`konteynerlar/xitoy_parsed/`ga yozadi) → `_main_py_ishga_tushir()`. `kont:yoq` → bekor.

### C) 🔄 Konteyner holatini o'zgartirish (`keldi_belgi`)
**Holati: to'liq ishlayapti.** `keldi_menu` ekran:
- **Yo'lda → Keldi**: sana/ISO so'raladi, fayl(lar) topiladi (bitta yoki checkbox-ko'p), `_D.xlsx` ga qayta nomlanadi, `kont_rasm.generate_kelgan_rasm` bilan rasm tayyorlanadi, guruhga yuborish so'raladi.
- **KELDI ni qaytarish**: `_D.xlsx` → `.xlsx`.
- **Guruhga yuborish:** caption matni kutiladi → `.env`dagi `KELGAN_YUKLAR_CHAT_ID`/`TOPIC_ID` guruhiga rasm(lar)+izoh.

---

## Sozlamalar

### Admin — `settings_kb`
- **🔄 Ma'lumotlarni yangilash**: `git pull origin main` → `main.py` subprocess → `xlsx_refresh(force=True)`. Ishlaydi. (DIQQAT: `.py` kod o'zgarishlari uchun bu yetarli emas — `pm2 restart`/`server_yangilash.bat` kerak.)
- **🌐 Tilni o'zgartirish**: ishlaydi.
- **🗑 Buyurtmani tozalash**: kanal tanlash → tasdiq → `buyurtma_{kanal}.json` o'chadi. Ishlaydi.
- **🇨🇳 Xitoy ostatkani tozalash**: xuddi shunday, `xitoy_{kanal}.json`. Ishlaydi.

### Foydalanuvchi (filial) — `settings_kb_user`
- **🌐 Tilni o'zgartirish**: ishlaydi.
- **📞 Admin bilan bog'lanish**: `.env`dagi `SUPPORT_PHONE`/`SUPPORT_USERNAME`. Ishlaydi.

---

## Boshqa inline / yordamchi callbacklar

| callback_data | Vazifa | Holati |
|---|---|---|
| `lang:cyr/lat`, `lang_pick` | Til tanlash | ishlaydi |
| `yolda_barchasi` | Yo'lda holat | ishlaydi |
| `kont_noop` | Bo'sh (placeholder) | ataylab hech narsa qilmaydi |
| `tozala_no` | Tozalashni bekor | ishlaydi |
| `kg_send/kg_cancel/kg_send_multi/kg_cancel_multi` | Guruhga yuborish boshqaruvi | ishlaydi |
| `kont_bekor:*` | Bekor qilindi | ishlaydi |

---

## Aniqlangan muammolar va o'lik kod (xulosa)

1. **JIDDIY — O'sh (🇰🇬) kanali buzuq (tasdiqlangan):** `config.py:114`da `CH_KEY` kalitida `" osh"` (ortiqcha bo'sh joy). "Buyurtma yig'ish → O'sh" yoki "Konteyner yuklash → O'sh" tanlanganda `CH_KEY["osh"]` → `KeyError`, bot javob bermay qoladi. **Tuzatish:** `config.py:114`da `" osh"` → `"osh"` (bitta qatorlik tuzatish).
2. **O'lik/eskirgan kod:** `action=="yolda"` → bo'sh `pass`; `action=="upload"` (tugmasiz); `ui.py:kamomat_ko_rish` (chaqirilmaydi); `texts.py:search_stub`, `kritik_stub`, `b_status`, `b_search`, `b_yolda`, `b_upload` — foydalanuvchiga ko'rinmaydi; `status_kb`/`status` ekrani ochilmaydi.
3. **Kichik nomuvofiqlik:** `zakaz_no`dagi fallback `["asosiy","cex","osh"]` — "cex" emas "sex" bo'lishi kerak (amaliy ta'siri yo'q).
4. **Ish papkasidagi `handlers.py`/`parsers.py` hozircha kesilgan holatda (git: `M`, uncommitted):** oxirgi ~30-40 qator yo'q, sintaktik xato — ish nusxasi ishga tushmaydi. To'g'ri versiya `git show HEAD:<fayl>` da (bu hujjat shu asosda tuzildi).

**Umumiy baho:** Botning asosiy oqimlari (Buyurtma Excel, Tasdiqlangan buyurtma, Yuklatish rejasi, Qidiruv/karta, Yo'ldagi yuklar, Yo'lga konteyner qo'shish, Keldi/Qaytarish holati, guruhga rasm, Sozlamalar, Yangilash) — **Asosiy va Tsex kanallari uchun to'liq ishlaydi**. Ikkita muhim kamchilik bor: (1) O'sh kanali `CH_KEY` xatosi, (2) ish papkasidagi ikki fayl hozircha kesilgan/commit qilinmagan holatda.
