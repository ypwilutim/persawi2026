# Persawi 2026 - Media Center

Aplikasi dokumentasi Photos & Video untuk Persawi 2026 (YPWI Luwu Timur).

## Stack

- Frontend statis: HTML + Tailwind CSS, di-deploy ke shared hosting (Apache/cPanel).
- Backend: Google Apps Script Web App + Google Spreadsheet + Google Drive.
- Penyimpanan file besar: Cloudflare R2 (upload langsung dari browser, backup ke server via rclone).

## Struktur

- `public.html` - galeri publik (tanpa login).
- `login.html` - halaman login.
- `register.html` - registrasi akun media (nama, asal sekolah, nomor telepon).
- `media.html` - panel upload dan daftar tugas untuk PJ Media tiap unit.
- `admin.html` - dashboard koordinator: kelola rundown, rekap file, tambah akun, pengaturan.
- `config.js` - URL API dan helper frontend.
- `code.gs` - backend Apps Script.
- `index.html` - redirect root ke `public.html` untuk shared hosting.

## Registrasi Media

Media dapat mendaftar sendiri melalui `register.html` dengan mengisi nama, asal sekolah, dan nomor telepon. Asal sekolah berupa dropdown yang diambil dari key `SEKOLAH` di sheet `Config` (format dipisah koma). Media tidak dapat menulis manual; jika sekolah belum ada, tersedia tautan untuk menghubungi koordinator agar sekolah ditambahkan ke daftar. Username dibuat dari nama depan + ID unik (contoh: `AKBAR_6`), password default `persawi2026`, dan password dapat diganti dari halaman `media.html`. Satu akun hanya untuk satu unit sekolah. Admin dapat mengubah daftar sekolah melalui form Konfigurasi Sistem dan melihat akun media di `admin.html`.

## Alur Tugas Media

1. Admin menambahkan rundown di `admin.html` (tanggal, jam, judul, kategori, lokasi, keterangan).
2. Setiap rundown menjadi tugas untuk seluruh unit media.
3. Media membuka `media.html`, memilih tugas, lalu mengunggah dokumentasi. Beberapa file dapat dipilih sekaligus dan setiap file menampilkan progres persentase.
4. Upload dengan tugas terpilih otomatis menandai tugas selesai untuk unit tersebut.
5. Admin melihat persentase penyelesaian per tugas dan total; media melihat persentase tugas unitnya.

Sheet `Rundown` menyimpan daftar acara. Sheet `TaskProgress` menyimpan status per unit. Kolom `task_id` di sheet `Media` menghubungkan file dengan tugas.

## Setup Apps Script

1. Buat Google Spreadsheet baru.
2. Extensions -> Apps Script, lalu tempel isi `code.gs`.
3. Jalankan fungsi `setupDatabaseSheets` sekali.
4. Di sheet `Config`, isi `FOLDER_ID` dengan ID folder Google Drive tujuan.
5. Jalankan fungsi `setupR2Config` sekali, lalu buka Project Settings -> Script Properties dan isi:
   - `R2_ACCOUNT_ID`
   - `R2_ACCESS_KEY_ID`
   - `R2_SECRET_ACCESS_KEY`
   - `R2_BUCKET`
   - `R2_PUBLIC_BASE` (contoh: `https://pub-xxxx.r2.dev`)
6. Deploy -> New deployment -> Web app.
   - Execute as: Me.
   - Who has access: Anyone.
7. Salin URL deployment (`.../exec`) ke `config.js` pada `API_URL`.
8. Setelah mengubah `code.gs`, deploy ulang dengan **New version** agar perubahan aktif.

Akun awal:

- Admin: `admin` / `admin2026`
- Media: `malili1` / `media2026`

## Deploy ke Shared Hosting

Tidak ada build step. Unggah file berikut ke `public_html` (atau folder web root shared hosting):

```text
public.html
login.html
register.html
media.html
admin.html
index.html
config.js
```

`index.html` mengarahkan root `/` ke `public.html`. Jika shared hosting memakai `DirectoryIndex`, cukup pastikan `index.html` atau `public.html` menjadi default. Tidak perlu PHP, Node.js, atau environment variable di shared hosting.

## Upload File Besar (Cloudflare R2)

File sampai `MAX_FILE_MB` (default 25 MB) tetap memakai jalur base64 + Google Drive. File di atas batas tersebut hingga `MAX_R2_MB` (default 2048 MB, maksimal 5120 MB) diunggah langsung dari browser ke Cloudflare R2 memakai presigned PUT URL. File tidak melewati shared hosting; Apps Script hanya membuat dan memverifikasi URL.

### Setup R2

1. Buat bucket R2 dan aktifkan akses publik (r2.dev atau custom domain) untuk pemutaran galeri.
2. Atur CORS bucket:
   - Allowed origin: domain shared hosting Anda, contoh `https://domain-anda.com`
   - Methods: `PUT`, `GET`, `HEAD`, `DELETE`
   - Allowed headers: `*`
3. Buat API token R2 dengan izin read/write objek.
4. Isi Script Properties Apps Script:
   - `R2_ACCOUNT_ID`
   - `R2_ACCESS_KEY_ID`
   - `R2_SECRET_ACCESS_KEY`
   - `R2_BUCKET`
   - `R2_PUBLIC_BASE` (contoh: `https://pub-xxxx.r2.dev`)
5. Deploy ulang Apps Script dengan **New version**.

Key objek berbentuk `<slug-unit>/YYYY/MM/DD/<timestamp>-<random>-<filename>`. Apps Script memverifikasi prefix unit, ukuran file, dan keberadaan objek lewat HEAD request sebelum menyimpan metadata. Kredensial R2 tidak pernah dikirim ke browser.

### Backup Server

```bash
rclone sync r2:persawi /backup/persawi/current \
  --backup-dir /backup/persawi/deleted \
  --checksum --fast-list --transfers 4 \
  --retries 10 --low-level-retries 20 --timeout 5m \
  --log-file /var/log/persawi-rclone.log --log-level INFO
```

Jalankan lewat systemd service + timer (`OnUnitActiveSec=1h`, `Persistent=true`). Objek yang dihapus dari R2 tetap tersimpan di direktori `deleted`.

## Catatan

- `MAX_FILE_MB` dibatasi maksimal 25 MB karena upload Drive dikirim sebagai base64 JSON. `MAX_R2_MB` (1-5120 MB) adalah batas upload langsung ke R2.
- Kolom `storage_type` di sheet `Media` bernilai `drive` atau `r2` (baris lama dianggap `drive`).
- Token sesi dibuat otomatis dan disimpan di Script Properties (`AUTH_SECRET`).
- Endpoint `addUser`, `deleteMedia`, `updateConfig`, `upload`, `completeR2Upload`, `getUploadUrl`, dan `deleteR2Object` memerlukan token sesi. `deleteR2Object` hanya untuk admin.
- Jangan pernah menaruh kredensial R2 di `config.js`, HTML, atau Git. Kredensial hanya di Script Properties Apps Script dan konfigurasi rclone server.
- Ganti password default `admin`/`admin2026` dan `malili1`/`media2026` di sheet `Users` sebelum dipakai.
- Jangan menaruh token GitHub di URL remote. Gunakan SSH, GitHub CLI, atau credential helper.
- Tugas dianggap selesai jika ada file yang diunggah dengan `task_id` terkait. Menghapus file terakhir untuk tugas tersebut akan mengembalikan status menjadi belum selesai.