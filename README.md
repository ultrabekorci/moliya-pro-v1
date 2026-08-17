# Moliya-Pro v2

Kichik va o'rta biznes uchun moliyaviy boshqaruv tizimi.
**TypeScript** + **Google Apps Script** + **Google Sheets** (ma'lumotlar bazasi sifatida).

Bu — [Moliya-Pro v1](https://github.com/ultrabekorci/moliya-pro-v1) ning to'liq qayta yozilgan
versiyasi. Xavfsizlik, hisob-kitob aniqligi va sozlanuvchanlik nuqtai nazaridan noldan qurilgan.

---

## Nima o'zgardi

| | v1 | v2 |
|---|---|---|
| **Autentifikatsiya** | Barcha loginlar va **ochiq matndagi parollar** HTML sahifaga joylashtirilardi; tekshiruv brauzerda bajarilardi | Parollar tuz (salt) bilan xeshlanadi (iterativ HMAC-SHA256); tekshiruv **faqat serverda**; sessiya tokeni |
| **Ruxsatlar** | Faqat UI'ni yashirardi; server hech narsani tekshirmasdi | Har bir RPC metodi uchun server tomonda majburiy tekshiruv, **fail-closed** |
| **To'lov turlari** | `{Naqd, P2P, Dollar, Bank}` kodga qotirib yozilgan — Uzcard/Humo tushumlari balansdan **tushib qolardi** | Ixtiyoriy sondagi **hisoblar**, har biri o'z valyutasida |
| **Firmalar** | `"Greenpen"` / `"Smartmiz"` kodda; firma savdo nuqtasi nomidan `includes("SMARTMIZ")` bilan topilardi | Firmalar — oddiy katalog yozuvi; nuqta → firma bog'lanishi `firmId` orqali |
| **Valyuta** | Dollar chiqimi UZS ga o'girib saqlanardi, asl summa **izohdan regexp bilan** ajratib olinardi; tahrirlashda summa qayta ko'paytirilardi | Summa har doim hisobning **o'z valyutasida**; UZS ekvivalenti alohida maydonda, faqat hisobot uchun |
| **Bir vaqtda ishlash** | `LockService` yo'q, `getLastRow()+1` poygasi | Barcha yozuv amallari qulf ostida; ID — UUID |
| **O'chirish** | Qator raqami bo'yicha; keyingi amallar noto'g'ri yozuvga tegardi | Barqaror UUID; tranzaksiyalar uchun **yumshoq o'chirish** + amallar tarixi |
| **XSS** | `innerHTML` ga foydalanuvchi matni to'g'ridan-to'g'ri qo'yilardi | DOM faqat `textContent` orqali quriladi |
| **Xatoliklar** | `withFailureHandler` deyarli yo'q — server xatosida ham "Muvaffaqiyatli!" chiqardi | Yagona `Promise` asosidagi RPC; xato e'tiborsiz qololmaydi |
| **Yuklash** | Har bir tab almashuvda butun jadval klientga tushardi | Server tomonda filtrlash va sahifalash |
| **Tashqi kutubxonalar** | Bootstrap, Chart.js, SweetAlert, Boxicons (ikki marta yuklangan CDN) | **Nol tashqi bog'liqlik** — o'z CSS, o'z SVG grafiklari |
| **Testlar** | Yo'q | 93 ta birlik testi (pul, sana, balans, statistika, ruxsatlar, validatsiya) |

---

## Arxitektura

```
src/
├── shared/      Server va klient BIRGA ishlatadigan sof kod (test bilan qoplangan)
│   ├── types.ts        Domen modeli
│   ├── money.ts        Butun sonli pul arifmetikasi (minor birlik)
│   ├── dates.ts        Vaqt mintaqasiga bog'liq bo'lmagan sana mantiqi
│   ├── permissions.ts  Ruxsatlar modeli (fail-closed)
│   ├── balance.ts      Balans hisoblash dvigateli
│   ├── stats.ts        Dashboard statistikasi
│   ├── validation.ts   Umumiy validatsiya qoidalari
│   └── api.ts          Klient ↔ server shartnomasi (tipizatsiyalangan)
│
├── server/      Google Apps Script tomoni
│   ├── main.ts         doGet / rpc / setup — kirish nuqtalari
│   ├── router.ts       YAGONA RPC eshigi + ruxsat tekshiruvi
│   ├── sheets.ts       Sheets ustidagi tipizatsiyalangan repozitoriy (+ LockService)
│   ├── db.ts           Jadval sxemalari
│   ├── migrate.ts      v1 → v2 ma'lumot ko'chirish
│   └── services/       auth, users, catalog, transactions, fx, config, audit
│
└── client/      Brauzer tomoni (framework'siz, tipizatsiyalangan)
    ├── main.ts         Qobiq, navigatsiya, sessiya
    ├── api.ts          Promise asosidagi RPC klienti
    ├── dom.ts          XSS'dan himoyalangan DOM yordamchilari
    ├── charts.ts       Qo'lda yozilgan SVG grafiklar
    ├── store.ts        Holat (faqat token saqlanadi, rol/ruxsat emas)
    └── views/          login, dashboard, transactions, settings
```

Ikkala tomon ham `shared/` dagi **bitta** implementatsiyadan foydalanadi, shuning uchun
klientdagi balans serverdagi balansdan farq qilishi mumkin emas.

---

## O'rnatish

### 1. Talablar

- Node.js 20+
- Google hisobi
- [clasp](https://github.com/google/clasp): `npm i -g @google/clasp && clasp login`

### 2. Loyihani tayyorlash

```bash
git clone <repo-url> moliya-pro
cd moliya-pro
npm install
npm run check      # typecheck + lint + test + build
```

### 3. Google Sheets va Apps Script

1. Yangi Google Sheets hujjati yarating.
2. **Kengaytmalar → Apps Script** ni oching.
3. Script ID ni nusxalang (**Loyiha sozlamalari** sahifasida).
4. `.clasp.json` yarating:

```bash
cp .clasp.json.example .clasp.json
# scriptId ni to'ldiring
```

### 4. Yuklash va ishga tushirish

```bash
npm run push       # build + clasp push
```

So'ng Apps Script muharririda:

1. `setup` funksiyasini bir marta ishga tushiring — varaqlar va boshlang'ich katalog yaratiladi.
2. `authorizeExternalRequests` ni ishga tushiring — valyuta kursi uchun internet ruxsati so'raladi.
3. **Deploy → New deployment → Web app** (`Execute as: Me`, `Who has access: Anyone`).
4. Chiqqan havolani oching — **birinchi administrator** yaratish oynasi ochiladi.

> Birinchi ishga tushirish oynasi faqat tizimda hech qanday foydalanuvchi
> bo'lmaganda ko'rinadi. Deploydan so'ng darhol administrator yarating.

### 5. Ixtiyoriy: sessiyalarni tozalash

Apps Script → **Triggers** → `purgeSessions` funksiyasiga kunlik trigger qo'shing.

---

## Sozlash (kodni o'zgartirmasdan)

Hamma narsa **Sozlamalar** bo'limidan boshqariladi:

- **Firmalar** — nechta bo'lsa shuncha. Har biriga umumiy xarajatlarni taqsimlash qoidasi:
  *daromadga proporsional*, *teng*, yoki *taqsimlanmasin*.
- **Hisoblar** — Naqd, Uzcard, Humo, Bank, Dollar kassa… Har biri o'z valyutasida (UZS, USD, EUR…).
- **Savdo nuqtalari** — firmaga biriktiriladi. `Daromadga kirmasin` bayrog'i bilan
  ichki kassa/direktor kabi nuqtalarni statistikadan chiqarish mumkin.
- **Kategoriyalar** — kirim / chiqim / ikkalasi.
- **Xodimlar** — 4 ta rol (administrator, menejer, operator, kuzatuvchi) va har bir
  bo'lim bo'yicha aniq huquqlar matritsasi.
- **Umumiy** — hisobot valyutasi, sessiya muddati, maksimal summa, kurs manbasi.

---

## v1 dan ma'lumot ko'chirish

Eski baza bilan bir xil Google Sheets hujjatida ishlayotgan bo'lsangiz:

```
Apps Script muharriri → migrateDryRun   (hech narsa yozmaydi, faqat hisobot beradi)
Apps Script muharriri → migrateFromV1   (haqiqiy ko'chirish)
```

Tafsilotlar: [`docs/MIGRATION.md`](docs/MIGRATION.md).

---

## Ishlab chiqish

```bash
npm run typecheck    # TypeScript tekshiruvi
npm run lint         # ESLint
npm run test         # Vitest
npm run test:watch   # kuzatuv rejimida
npm run coverage     # qamrov hisoboti
npm run build        # dist/ ni yig'ish
npm run check        # hammasi birga
npm run push         # build + clasp push
```

Ma'lumotlar bazasi strukturasi: [`docs/SCHEMA.md`](docs/SCHEMA.md).
Xavfsizlik yechimlari: [`docs/SECURITY.md`](docs/SECURITY.md).

---

## Litsenziya

[MIT](LICENSE)
