# Persawi 2026 - Media Center

Aplikasi dokumentasi Photos & Video untuk Persawi 2026 (YPWI Luwu Timur).

## Stack

- Frontend statis: HTML + Tailwind CSS, di-deploy ke Vercel.
- Backend: Google Apps Script Web App + Google Spreadsheet + Google Drive.

## Struktur

- `public.html` - galeri publik (tanpa login).
- `login.html` - halaman login.
- `media.html` - panel upload untuk PJ Media tiap unit.
- `admin.html` - dashboard koordinator: rekap file, tambah akun, pengaturan.
- `config.js` - URL API dan helper frontend.
- `code.gs` - backend Apps Script.
- `vercel.json` - rewrite root ke `public.html`.

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