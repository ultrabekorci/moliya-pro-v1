# Ma'lumotlar bazasi strukturasi

Har bir varaqning birinchi qatori — sarlavha. Kod ustunlarni **nom bo'yicha** topadi,
shuning uchun ustunlar tartibini o'zgartirish yoki oraga yangi ustun qo'shish
dasturni buzmaydi. Yetishmayotgan varaq/ustun ilova ishga tushganda avtomatik yaratiladi.

Barcha yozuvlar barqaror **UUID** `id` ga ega. Qator raqami hech qachon
identifikator sifatida ishlatilmaydi.

---

## `Config`

Kalit-qiymat juftliklari.

| Ustun | Izoh |
|---|---|
| `key` | Sozlama nomi (`baseCurrency`, `organizationName`, …) |
| `value` | Qiymat (matn ko'rinishida) |

Ma'lum kalitlar: `baseCurrency`, `organizationName`, `sessionTtlMinutes`,
`maxLoginAttempts`, `loginLockoutMinutes`, `maxTransactionAmount`, `maxNoteLength`,
`sharedExpenseAllocation`, `fxProvider`.

---

## `Firms`

| Ustun | Tur | Izoh |
|---|---|---|
| `id` | UUID | |
| `name` | matn | Firma nomi |
| `active` | mantiqiy | Nofaol firma yangi yozuvlarda tanlanmaydi |
| `expenseAllocation` | `direct` \| `proRataIncome` \| `equal` | Umumiy xarajatlarni taqsimlash usuli |
| `sortOrder` | son | Ko'rsatish tartibi |

---

## `Accounts` — hisoblar (to'lov turlari)

| Ustun | Tur | Izoh |
|---|---|---|
| `id` | UUID | |
| `name` | matn | Naqd, Uzcard, Bank hisobi, Dollar kassa… |
| `currency` | ISO 4217 | `UZS`, `USD`, `EUR`… |
| `active` | mantiqiy | |
| `showInBalance` | mantiqiy | Balans panelida ko'rinsinmi |
| `sortOrder` | son | |

> Hisob valyutasini o'zgartirish faqat unda yozuvlar bo'lmaganda mumkin —
> aks holda mavjud qoldiq ma'nosini yo'qotadi.

---

## `Points` — savdo nuqtalari

| Ustun | Tur | Izoh |
|---|---|---|
| `id` | UUID | |
| `name` | matn | |
| `firmId` | UUID | `Firms.id` ga havola |
| `active` | mantiqiy | |
| `excludeFromRevenue` | mantiqiy | `true` — tushumi daromad statistikasiga kirmaydi |
| `sortOrder` | son | |

---

## `Categories`

| Ustun | Tur | Izoh |
|---|---|---|
| `id` | UUID | |
| `name` | matn | |
| `kind` | `income` \| `expense` \| `both` | Qaysi formada tanlanadi |
| `active` | mantiqiy | |
| `sortOrder` | son | |

---

## `Users`

| Ustun | Tur | Izoh |
|---|---|---|
| `id` | UUID | |
| `login` | matn | Kichik harflarda, takrorlanmas |
| `displayName` | matn | |
| `passwordHash` | base64 | Iterativ HMAC-SHA256 natijasi |
| `passwordSalt` | matn | Har bir foydalanuvchi uchun alohida |
| `passwordIterations` | son | Iteratsiyalar soni (kelajakda oshirish uchun) |
| `role` | `admin` \| `manager` \| `operator` \| `viewer` | |
| `status` | `active` \| `blocked` | |
| `permissions` | JSON | Rol standartiga qo'shimcha huquqlar |
| `createdAt` / `updatedAt` / `lastLoginAt` | ISO 8601 | |

> **Parol hech qachon ochiq matnda saqlanmaydi** va hech qachon klientga yuborilmaydi.

---

## `Transactions`

| Ustun | Tur | Izoh |
|---|---|---|
| `id` | UUID | Bitta qator |
| `batchId` | UUID | Bir amalda kiritilgan qatorlarni bog'laydi |
| `date` | `YYYY-MM-DD` | Matn sifatida saqlanadi |
| `kind` | `income` \| `expense` \| `transfer` | |
| `firmId` | UUID \| bo'sh | Odatda nuqta orqali aniqlanadi |
| `pointId` | UUID \| bo'sh | |
| `categoryId` | UUID \| bo'sh | Chiqim uchun majburiy |
| `accountId` | UUID | Kirim: pul kirgan hisob. Chiqim/o'tkazma: pul chiqqan hisob |
| `amountMinor` | butun son | **`accountId` hisobining o'z valyutasida**, minor birlikda |
| `currency` | ISO 4217 | `accountId` valyutasi |
| `counterAccountId` | UUID \| bo'sh | Faqat o'tkazma: pul tushgan hisob |
| `counterAmountMinor` | butun son \| bo'sh | Tushgan summa (o'z valyutasida) |
| `counterCurrency` | ISO 4217 \| bo'sh | |
| `fxRate` | son | Hisobot valyutasiga o'tkazish kursi |
| `baseAmountMinor` | butun son | Hisobot valyutasidagi ekvivalent — **faqat statistika uchun** |
| `note` | matn | |
| `createdAt` / `createdBy` | | |
| `updatedAt` / `updatedBy` | | |
| `deletedAt` / `deletedBy` | | Yumshoq o'chirish |

### Balans qoidasi

```
income   →  accountId          += amountMinor
expense  →  accountId          -= amountMinor
transfer →  accountId          -= amountMinor
            counterAccountId   += counterAmountMinor
```

`deletedAt` to'ldirilgan qatorlar hisob-kitobga umuman kirmaydi.

### Nega summa hisobning o'z valyutasida?

v1 da dollar chiqimi UZS ga o'girib saqlanardi va asl summa izohga
`(Aslida: $500 @ 12500)` ko'rinishida yozilardi. Natijada:

- balans izohni regexp bilan tahlil qilishga majbur edi;
- yozuvni tahrirlaganda summa **yana bir marta** kursga ko'paytirilardi.

v2 da `amountMinor` — bu har doim o'sha hisobning haqiqiy valyutasidagi summa.
`baseAmountMinor` faqat hisobotlarni umumlashtirish uchun; u balansga ta'sir qilmaydi.

---

## `FxRates`

| Ustun | Tur | Izoh |
|---|---|---|
| `id` | UUID | |
| `date` | `YYYY-MM-DD` | Kurs qaysi kunga tegishli |
| `currency` | ISO 4217 | |
| `rate` | son | 1 birlik = necha hisobot valyutasi |
| `source` | matn | `cbu.uz`, `manual`, `base` |
| `fetchedAt` | ISO 8601 | |

Kurs topilmasa, eng yaqin oldingi sanadagi kurs ishlatiladi. Internet ishlamasa
kursni Sozlamalar → Valyuta kursi bo'limidan qo'lda kiritish mumkin.

---

## `AuditLog`

| Ustun | Izoh |
|---|---|
| `id`, `at` | |
| `userId`, `userLogin` | Kim bajardi |
| `action` | `transaction.create`, `user.delete`, `config.save`… |
| `entity`, `entityId` | Nimaga tegdi |
| `details` | JSON qisqartma (900 belgidan oshmaydi) |
