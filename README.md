# Persawi 2026 - Media Center

Aplikasi dokumentasi Photos & Video untuk Persawi 2026 (YPWI Luwu Timur).

## Stack

- Frontend statis: HTML + Tailwind CSS, di-deploy ke Vercel.
- Backend: Google Apps Script Web App + Google Spreadsheet + Google Drive.

## Struktur

- `public.html` - galeri publik (tanpa login).
- `login.html` - halaman login.
- `register.html` - registrasi akun media (nama, asal sekolah, nomor telepon).
- `media.html` - panel upload dan daftar tugas untuk PJ Media tiap unit.
- `admin.html` - dashboard koordinator: kelola rundown, rekap file, tambah akun, pengaturan.
- `config.js` - URL API dan helper frontend.
- `code.gs` - backend Apps Script.
- `vercel.json` - rewrite root ke `public.html`.

## Registrasi Media

Media dapat mendaftar sendiri melalui `register.html` dengan mengisi nama, asal sekolah, dan nomor telepon. Asal sekolah berupa dropdown yang diambil dari key `SEKOLAH` di sheet `Config` (format dipisah koma). Media tidak dapat menulis manual; jika sekolah belum ada, tersedia tautan untuk menghubungi koordinator agar sekolah ditambahkan ke daftar. Username dibuat dari nama depan + ID unik (contoh: `AKBAR_6`), password default `persawi2026`, dan password dapat diganti dari halaman `media.html`. Satu akun hanya untuk satu unit sekolah. Admin dapat mengubah daftar sekolah melalui form Konfigurasi Sistem dan melihat akun media di `admin.html`.

## Alur Tugas Media

1. Admin menambahkan rundown di `admin.html` (tanggal, jam, judul, kategori, lokasi, keterangan).
2. Setiap rundown menjadi tugas untuk seluruh unit media.
3. Media membuka `media.html`, memilih tugas, lalu mengunggah dokumentasi.
4. Upload dengan tugas terpilih otomatis menandai tugas selesai untuk unit tersebut.
5. Admin melihat persentase penyelesaian per tugas dan total; media melihat persentase tugas unitnya.

Sheet `Rundown` menyimpan daftar acara. Sheet `TaskProgress` menyimpan status per unit. Kolom `task_id` di sheet `Media` menghubungkan file dengan tugas.

## Setup Apps Script

1. Buat Google Spreadsheet baru.
2. Extensions -> Apps Script, lalu tempel isi `code.gs`.
3. Jalankan fungsi `setupDatabaseSheets` sekali.
4. Di sheet `Config`, isi `FOLDER_ID` dengan ID folder Google Drive tujuan.
5. Deploy -> New deployment -> Web app.
   - Execute as: Me.
   - Who has access: Anyone.
6. Salin URL deployment (`.../exec`) ke `config.js` pada `API_URL`.
7. Setelah mengubah `code.gs`, deploy ulang dengan **New version** agar perubahan aktif.

Akun awal:

- Admin: `admin` / `admin2026`
- Media: `malili1` / `media2026`

## Deploy Vercel

```bash
git push
```

Vercel akan otomatis deploy dari branch `main`. Root `/` diarahkan ke `public.html` melalui `vercel.json`.

## Catatan

- `MAX_FILE_MB` dibatasi maksimal 25 MB karena upload dikirim sebagai base64 JSON.
- Token sesi dibuat otomatis dan disimpan di Script Properties (`AUTH_SECRET`).
- Endpoint `addUser`, `deleteMedia`, `updateConfig`, dan `upload` memerlukan token sesi.
- Ganti password default `admin`/`admin2026` dan `malili1`/`media2026` di sheet `Users` sebelum dipakai.
- Jangan menaruh token GitHub di URL remote. Gunakan SSH, GitHub CLI, atau credential helper.
- Tugas dianggap selesai jika ada file yang diunggah dengan `task_id` terkait. Menghapus file terakhir untuk tugas tersebut akan mengembalikan status menjadi belum selesai.