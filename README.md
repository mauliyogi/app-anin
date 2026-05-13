# SPS Transfer Antar Gudang

Aplikasi web statis untuk operasional transfer antar gudang, rekap form, ekspor Excel/PDF, file import SPS, dan pemantauan inventori produk.

## Cara menjalankan lokal

```bash
python3 -m http.server 8000
```

Buka `http://127.0.0.1:8000/`.

## Cara deploy

Deploy isi repository ini ke static hosting apa pun (Netlify, Vercel static, GitHub Pages, Nginx, Apache, atau object storage). Entry point produksi adalah `index.html`.

## Catatan operasional produksi

- Data tersimpan di browser pengguna melalui `localStorage` dengan key `sps_state`.
- Ekspor Excel membutuhkan koneksi ke CDN SheetJS/XLSX yang dideklarasikan di `index.html`.
- Untuk operasi multi-user, gunakan prosedur backup/export rutin atau integrasikan storage backend sebelum dipakai bersama banyak perangkat.
- Halaman **Inventori** menghitung estimasi stok gudang dengan rumus: `stok awal - total naik + total turun`.

## Pemeriksaan rilis dasar

Sebelum deploy, jalankan:

```bash
node --check /tmp/app-script.js
python3 -m http.server 8000 --bind 127.0.0.1
curl -I http://127.0.0.1:8000/
```
