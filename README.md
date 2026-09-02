# Dashboard Progres Pemenuhan Headcount

Dashboard React (Vite) buat pantau progres pemenuhan Headcount level staff (N6),
datanya disimpan di Firestore (real-time, dibagi ke semua orang yang buka link-nya)
dan di-hosting gratis lewat GitHub Pages.

Lihat langkah-langkah setup lengkap di percakapan Claude, atau ringkasannya di bawah.

## Coba di lokal dulu

```bash
npm install
npm run dev
```

Buka `http://localhost:5173` — masih pakai data contoh sampai `src/firebase.js` diisi
config asli dan di-deploy.

## Yang wajib diubah sebelum deploy

1. **`src/firebase.js`** — ganti semua `GANTI_DENGAN_...` dengan config Firebase
   project lo sendiri (Firebase Console → Project settings → Your apps).
2. **`vite.config.js`** — ganti `base: '/headcount-dashboard/'` sesuai nama repo
   GitHub lo persis (huruf besar/kecil dan tanda hubung harus sama).
3. **Firestore Rules** — copot isi `firestore.rules` ke tab Rules di Firestore
   Console, klik Publish.

## Build manual (kalau gak pakai GitHub Actions)

```bash
npm run build
```

Hasilnya ada di folder `dist/`, tinggal upload ke hosting mana aja.

## Struktur data di Firestore

Dokumen: `headcount/current`

```json
{
  "records": [ ... baris dari sheet Headcount ... ],
  "candidates": [ ... baris dari sheet Kandidat ... ],
  "fileName": "Book9.xlsx",
  "updatedAt": "2026-09-02T08:00:00.000Z",
  "hasCandidateSheet": true
}
```

Setiap kali ada yang klik "Upload data baru" di web, dokumen ini di-replace
seluruhnya, dan semua orang yang lagi buka dashboard otomatis ke-update
(real-time lewat `onSnapshot`).

## Catatan keamanan

Firestore rules di project ini dibiarkan terbuka (siapa aja bisa baca DAN
tulis), sesuai yang diminta waktu setup. Artinya siapa pun yang punya link
bisa ganti data dashboard. Kalau nanti mau dibatasi cuma tim tertentu yang
boleh upload, lihat komentar contoh di `firestore.rules`.
