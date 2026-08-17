# v1 → v2 ma'lumot ko'chirish

Ko'chirish **eski varaqlarni o'zgartirmaydi** — ular joyida qoladi va istalgan
paytda solishtirish mumkin.

---

## Tartib

1. **Zaxira nusxa oling.** Google Sheets → *Fayl → Nusxa yaratish*. Bu qadamni
   o'tkazib yubormang.
2. v2 kodini o'sha jadvalga yuklang (`npm run push`).
3. Apps Script muharririda `setup` funksiyasini ishga tushiring.
4. **`migrateDryRun`** ni ishga tushiring. Hech narsa yozilmaydi — faqat hisobot
   qaytariladi:

```json
{
  "dryRun": true,
  "firms": 1,
  "accounts": 6,
  "points": 12,
  "categories": 9,
  "users": 4,
  "transactions": 3184,
  "skipped": ["482-qator: «Perechislenie» hisobi topilmadi"],
  "warnings": ["Barcha savdo nuqtalari bitta firmaga bog‘landi…"]
}
```

5. `skipped` ro'yxatini ko'rib chiqing. Agar biror narsa noto'g'ri ko'chayotgan
   bo'lsa, avval eski jadvalni tuzating va `migrateDryRun` ni qayta ishga tushiring.
6. Hammasi joyida bo'lsa — **`migrateFromV1`** ni ishga tushiring.
7. Veb-ilovani oching va balanslarni eski tizim ko'rsatgan qiymatlar bilan solishtiring.

> `migrateFromV1` yangi `Transactions` varag'i bo'sh bo'lmasa ishlamaydi —
> takroriy ko'chirish natijasida yozuvlar ikkilanib ketmaydi.

---

## Nima nimaga aylanadi

### `Data` varag'i → kataloglar

| Eski | Yangi |
|---|---|
| A ustun (firmalar) | `Firms` — birinchi nomi olinadi, bo'sh bo'lsa «Asosiy firma» |
| B ustun (savdo nuqtalari) | `Points` (hammasi bitta firmaga bog'lanadi) |
| C ustun (to'lov turlari) | `Accounts` |
| D ustun (kategoriyalar) | `Categories` |

Kodda qotirib yozilgan `Naqd / P2P / Bank / Dollar` to'lov turlari ham qo'shiladi.

**Valyuta taxmini:** hisob nomida `DOLLAR`, `USD` yoki `$` bo'lsa — `USD`;
`EVRO`/`EUR` bo'lsa — `EUR`; aks holda hisobot valyutasi. Ko'chirishdan so'ng
Sozlamalar → Hisoblar bo'limidan tekshiring.

**`excludeFromRevenue`:** v1 da kodga qotirib yozilgan ro'yxat
(`DONIYOR AKA`, `BOSHQA`, `DIREKTOR`, `KASSA`) mos nuqtalarga bayroq sifatida qo'yiladi.

### `Users` varag'i → `Users`

| Eski | Yangi |
|---|---|
| A — login | `login` (kichik harflarga keltiriladi) |
| B — parol (ochiq matn) | `passwordHash` + `passwordSalt` |
| C — rol | `admin` yoki `operator` |
| D — holat | `active` / `blocked` |
| E — huquqlar JSON | `permissions` (noma'lum kalitlar tashlanadi) |

> Eski parollar ishlashda davom etadi, lekin ular allaqachon oshkor bo'lgan
> bo'lishi mumkin. Ko'chirishdan so'ng barcha xodimlarga parolni almashtiring.

### `Kirim Chiqim` varag'i → `Transactions`

| Eski ustun | Yangi |
|---|---|
| A — sana | `date` (jadval vaqt mintaqasi bo'yicha `YYYY-MM-DD` ga keltiriladi) |
| C — savdo nuqtasi | `pointId` (+ `firmId` nuqta orqali) |
| D — kategoriya | `categoryId`; `O'tkazma`/`Transfer` → `kind: transfer` |
| E — to'lov turi | `accountId` |
| F — kirim | `kind: income`, `amountMinor` |
| G — chiqim | `kind: expense`, `amountMinor` |
| H — izoh | `note` |
| I — ID | `batchId` |
| J, K | O'tkazma uchun manba/manzil hisoblari |

**Dollar chiqimlari.** v1 da summa UZS ga o'girib saqlanar, asl USD qiymati
J ustunida yoki izohda `(Aslida: $500 @ 12500)` ko'rinishida bo'lardi.
Ko'chirishda:

- J ustunida son bo'lsa — o'sha olinadi;
- bo'lmasa izohdan `Aslida: $X` ajratib olinadi;
- ikkalasi ham topilmasa qator **o'tkazib yuboriladi** va `skipped` ga tushadi —
  bunday yozuvlarni qo'lda kiritish kerak.

`fxRate` = `UZS summa ÷ USD summa` sifatida qayta tiklanadi, `baseAmountMinor`
esa eski UZS qiymatiga teng bo'ladi — ya'ni tarixiy hisobotlar o'zgarmaydi.

---

## Ko'chirishdan keyin nima qilish kerak

1. **Savdo nuqtalarini firmalarga taqsimlang.** Hammasi bitta firmaga
   bog'langan holda keladi (v1 da firma savdo nuqtasi nomidan taxmin qilinardi).
   Sozlamalar → Firmalar bo'limidan kerakli firmalarni yarating, so'ng
   Savdo nuqtalari bo'limidan har birini o'z firmasiga biriktiring.
2. **Hisoblar valyutasini tekshiring.**
3. **Balanslarni solishtiring** — eski tizim ko'rsatgan qiymatlar bilan.
   Farq bo'lsa, sababi odatda `skipped` ro'yxatida ko'rinadi.
4. **Xodimlar parolini yangilang.**
5. Hammasi joyida bo'lsa, eski varaqlarni (`Kirim Chiqim`, `Data`) alohida
   hujjatga ko'chirib, asosiy jadvaldan olib tashlashingiz mumkin.
