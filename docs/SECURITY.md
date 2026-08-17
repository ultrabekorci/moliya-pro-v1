# Xavfsizlik

Bu hujjat v1 dagi zaifliklarni va v2 da ular qanday yopilganini tavsiflaydi.

---

## 1. Parollar sahifaga joylashtirilardi

**v1.** `doGet` barcha foydalanuvchilarni, jumladan **ochiq matndagi parollarni**,
HTML sahifaga yozib yuborardi:

```js
template.usersData = JSON.stringify(getAllUsers());   // Code.gs
var PRELOADED_USERS = <?!= usersData ?>;              // index.html
```

Veb-ilova havolasini bilgan har qanday odam sahifa manbasini ochib barcha
login/parollarni ko'ra olardi. `getAllUsers()` ham himoyasiz global funksiya edi.

**v2.** Sahifa butunlay statik — unda hech qanday ma'lumot yo'q. Foydalanuvchi
ma'lumoti faqat autentifikatsiyadan o'tgan `rpc` chaqiruvi orqali keladi va
`PublicUser` turida parol maydonlari umuman mavjud emas.

---

## 2. Autentifikatsiya klientda edi

**v1.** Tekshiruv brauzerda bajarilardi:

```js
var foundUser = LOCAL_USERS_DB.find((user) => user.u === u && user.p === p);
```

Server funksiyalari (`saveTransaction`, `deleteUser`, …) hech qanday tekshiruvsiz
`google.script.run` orqali chaqirilardi.

**v2.**

- Yagona kirish nuqtasi: `rpc(request)` → `router.dispatch()`.
- Har bir metod uchun kerakli ruxsat `GUARDS` jadvalida e'lon qilingan.
  Ro'yxatga kiritilmagan metod umuman chaqirilmaydi.
- Har bir chaqiruv sessiya tokeni bilan tekshiriladi.

---

## 3. Parollar ochiq matnda saqlanardi

**v2.** Parol tuz (salt) bilan **iterativ HMAC-SHA256** orqali xeshlanadi
(standart: 12 000 iteratsiya, foydalanuvchi bo'yicha saqlanadi — kelajakda oshirish mumkin).
Solishtirish vaqt bo'yicha doimiy (`constantTimeEquals`), shuning uchun javob
vaqtidan parolni tuslash mumkin emas.

Apps Script'da tayyor PBKDF2 yo'q, shuning uchun takrorlash qo'lda bajariladi.
Bu bitta SHA'dan ancha qimmatroq va lug'at hujumini sezilarli sekinlashtiradi.

---

## 4. Ruxsatlar fail-open edi

**v1.**

```js
if (!permissions || permissions.admin === true) { /* hammasi ochiq */ }
```

Huquqlar yuklanmagan bo'lsa foydalanuvchi **to'liq admin huquqini** olardi.

**v2.** `can()` funksiyasi fail-closed: ruxsat aniq `true` deb belgilanmagan
bo'lsa — yo'q. Test bilan qoplangan (`tests/permissions.test.ts`).

Bundan tashqari:

- ruxsat tekshiruvi serverda; UI faqat tugmalarni yashiradi;
- `tx.list` / `tx.listGrouped` natijasi foydalanuvchi ko'ra oladigan turlar
  bilan **serverda** cheklanadi — klient qanday filtr yuborishidan qat'i nazar;
- `admin` rolini shaxsiy sozlama bilan cheklab bo'lmaydi (o'zini qulflab
  qo'yishning oldi olinadi);
- oxirgi faol administratorni o'chirish yoki bloklash taqiqlangan.

---

## 5. Rolni brauzerda soxtalashtirish

**v1.** Rol va huquqlar `sessionStorage` da saqlanardi:

```js
sessionStorage.setItem('moliya_role', foundUser.r);
```

DevTools'da uni `Admin` qilib qo'yish kifoya edi.

**v2.** `sessionStorage` da **faqat token** saqlanadi. Rol va ruxsatlar har safar
serverdan olinadi (`auth.session`). Brauzerdagi qiymatga hech qachon ishonilmaydi.

---

## 6. XSS

**v1.** Jadvallar qator matn sifatida qurilardi:

```js
html += `<td>${item.note}</td>`;
```

Izohga `<img src=x onerror=...>` yozgan xodim boshqalarning brauzerida kod ishga
tushira olardi.

**v2.** DOM `document.createElement` + `textContent` orqali quriladi
(`src/client/dom.ts`). Foydalanuvchi matni hech qachon HTML sifatida
talqin qilinmaydi.

---

## 7. Clickjacking

**v1.** `setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)` —
sahifani istalgan saytga `<iframe>` sifatida joylash mumkin edi.

**v2.** `XFrameOptionsMode.DEFAULT` ishlatiladi.

---

## 8. Login urinishlarini cheklash

**v1.** Cheklov yo'q edi — parolni cheksiz tanlash mumkin edi.

**v2.** Ketma-ket noto'g'ri urinishlar `CacheService` da sanaladi; chegaradan
(standart: 5) oshsa login belgilangan muddatga (standart: 15 daqiqa) bloklanadi.
Ikkala qiymat ham sozlamalardan boshqariladi.

Foydalanuvchi topilmaganda ham parol tekshiruviga vaqt sarflanadi — mavjud
loginlarni javob vaqti bo'yicha aniqlash qiyinlashadi.

---

## 9. Sessiyalar

- Token — ikkita UUID birikmasi (128 bitdan ortiq entropiya).
- Saqlashda tokenning **SHA-256 xeshi** kalit sifatida ishlatiladi: Script
  Properties tarkibi sizib chiqsa ham tayyor token qo'lga tushmaydi.
- Muddati sozlamadan boshqariladi (standart: 8 soat).
- Parol o'zgarganda, foydalanuvchi bloklanganda yoki logini o'zgarganda
  uning **barcha sessiyalari** bekor qilinadi.
- `purgeSessions` funksiyasi muddati o'tganlarini tozalaydi (kunlik trigger).

---

## 10. Ma'lumot yaxlitligi

- Barcha yozuv amallari `LockService` ostida — bir vaqtda yozishdan kelib
  chiqadigan yo'qotishlar yo'q.
- Identifikatorlar — `Utilities.getUuid()`. v1 dagi `Date.now()` bir
  millisekundda ikki marta chaqirilsa to'qnashardi va begona yozuv o'chib ketardi.
- Tranzaksiyalar yumshoq o'chiriladi; kim va qachon o'chirgani `AuditLog` da qoladi.
- Ma'lumotnoma yozuvini o'chirish, agar u tranzaksiyalarda ishlatilgan bo'lsa,
  taqiqlanadi — o'rniga "nofaol" qilish taklif etiladi (yetim ma'lumot paydo bo'lmaydi).

---

## Nima himoyalanmaydi

Bu ro'yxat ataylab ochiq:

- **Google hisobi xavfsizligi** sizning zimmangizda. Skript egasining hisobi
  buzilsa, butun baza ochiladi. Ikki bosqichli tasdiqlashni yoqing.
- **Jadvalga to'g'ridan-to'g'ri kirish.** Google Sheets hujjatiga ulashilgan
  har qanday odam ma'lumotni to'g'ridan-to'g'ri o'zgartira oladi — ilova
  ruxsatlari bunga to'sqinlik qilmaydi. Hujjatni faqat o'zingizda qoldiring.
- **Deploy sozlamasi.** `Who has access: Anyone` bo'lsa, havolani bilgan har kim
  login sahifasini ko'radi (lekin ichkariga parolsiz kira olmaydi).
- **Ochiq matndagi eski parollar.** v1 dan ko'chirilgan parollar xeshlanadi,
  lekin ular allaqachon oshkor bo'lgan bo'lishi mumkin — ko'chirishdan so'ng
  barcha xodimlarga parolni almashtirishni tavsiya qiling.

## Zaiflik topsangiz

Muammoni ommaviy issue sifatida ochmang — repozitoriy egasiga to'g'ridan-to'g'ri yozing.
