# Task Management

Aplikasi manajemen tugas berbasis **Google Apps Script** dengan **Google Sheets** sebagai penyimpanan data dan **Google Drive** sebagai penyimpanan lampiran. Antarmuka web menyediakan dashboard, pengelolaan tugas, unggah massal, komentar, serta pencatatan dan tindak lanjut kendala.

## Fitur utama

- Dashboard ringkasan tugas: total, *To Do*, *In Progress*, *Done*, terlambat, prioritas mendesak, dan total jumlah data.
- Tambah, lihat, ubah, hapus lunak, dan pembaruan status beberapa tugas sekaligus.
- Filter dan pencarian berdasarkan status, prioritas, kategori, penanggung jawab, pembuat, dan kata kunci.
- Pengurutan data dan paginasi pada daftar tugas dan kendala.
- Impor tugas secara massal melalui template Excel, maksimal 1.000 baris per proses.
- Komentar dan balasan komentar pada setiap tugas.
- Pencatatan kendala, penyebab, status tindak lanjut, statistik, serta pemfilteran kendala.
- Unggah lampiran ke Google Drive dan penghapusan file lampiran.
- Identifikasi pengguna berdasarkan akun Google dan data pada sheet `USERS`.
- Cache selama 300 detik untuk mempercepat pemuatan pengguna, konfigurasi, dan data dashboard.
- Tampilan responsif menggunakan Bootstrap, Font Awesome, dan SheetJS.

## Struktur repository

| File | Keterangan |
|---|---|
| `Code.gs` | Backend Google Apps Script: akses spreadsheet, autentikasi pengguna, CRUD tugas, komentar, kendala, statistik, bulk upload, cache, dan lampiran Drive. |
| `index.html` | Antarmuka web, modal, tabel, filter, pagination, dan pemanggilan fungsi backend melalui `google.script.run`. |
| `README.md` | Dokumentasi instalasi, struktur data, penggunaan, dan pemeliharaan aplikasi. |

> Nama `index.html` harus dipertahankan karena fungsi `doGet()` memanggil `HtmlService.createTemplateFromFile("index")`.

## Teknologi

- Google Apps Script V8
- Google Sheets
- Google Drive
- HTML, CSS, dan JavaScript
- Bootstrap 5.3.0-alpha1
- Font Awesome 6.2.1
- SheetJS 0.18.5 untuk membaca berkas Excel di browser

Library antarmuka dimuat melalui CDN di `index.html`; tidak diperlukan `npm install` atau proses build.

## Persiapan Google Sheets

Spreadsheet yang digunakan ditentukan oleh `CONFIG.SPREADSHEET_ID` di `Code.gs`. Siapkan sheet berikut dengan nama yang sama persis.

### 1. `TRACKER`

Sheet utama untuk data tugas. Header yang digunakan aplikasi:

| Kolom | Kegunaan |
|---|---|
| `TASK_ID` | ID unik yang dibuat otomatis. |
| `TIMESTAMP` | Waktu pembuatan tugas. |
| `JUDUL` | Judul tugas. |
| `KATEGORI` | Kategori dari sheet `CONFIG`. |
| `DESKRIPSI` | Uraian tugas. |
| `JUMLAH_DATA` | Jumlah data/volume pekerjaan. |
| `ASSIGN_TO` | Nama penanggung jawab. |
| `ASSIGN_USER_ID` | ID pengguna yang ditugaskan. |
| `STATUS` | Status tugas dari sheet `CONFIG`. |
| `PRIORITY` | Prioritas dari sheet `CONFIG`. |
| `DUE_DATE` | Tenggat tugas. |
| `ATTACHMENT_LINK` | URL lampiran Google Drive atau tautan eksternal. |
| `CREATED_BY` | Email pembuat tugas. |
| `LAST_STATUS` | Status sebelumnya untuk kebutuhan pelacakan. |
| `LAST_UPDATE` | Waktu perubahan terakhir. |
| `IS_DELETED` | Penanda hapus lunak. |

### 2. `USERS`

| Kolom | Kegunaan |
|---|---|
| `USER_ID` | ID unik pengguna. |
| `NAMA_LENGKAP` | Nama yang tampil di aplikasi. |
| `ROLE` | Peran pengguna; nilai bawaan aplikasi adalah `user`. |
| `EMAIL` | Email akun Google untuk mencocokkan pengguna aktif. |

Pengguna yang menjalankan bulk upload wajib memiliki `USER_ID` dan email yang terdaftar pada sheet ini.

### 3. `COMMENTS`

| Kolom | Kegunaan |
|---|---|
| `COMMENT_ID` | ID unik komentar. |
| `TASK_ID` | Relasi ke tugas pada `TRACKER`. |
| `USER_ID` | Pengguna yang menulis komentar. |
| `COMMENT` | Isi komentar. |
| `TIMESTAMP` | Waktu komentar dibuat. |
| `PARENT_COMMENT_ID` | ID komentar induk untuk balasan; boleh kosong. |

### 4. `FOLLOW_UP_KENDALA`

| Kolom | Kegunaan |
|---|---|
| `ID` | ID unik catatan kendala. |
| `TASK_ID` | Relasi ke tugas. |
| `KENDALA` | Deskripsi kendala. |
| `PENYEBAB` | Penyebab kendala dari sheet `CONFIG`. |
| `STATUS` | Status tindak lanjut. |
| `CREATED_AT` | Waktu kendala dibuat. |
| `CREATED_BY` | Pengguna pembuat kendala. |
| `LAST_UPDATE` | Waktu pembaruan terakhir. |

### 5. `CONFIG`

Sheet ini menjadi sumber pilihan dropdown. Gunakan struktur berikut:

| Kolom | Isi |
|---|---|
| Kolom A | Tipe konfigurasi. |
| Kolom B | Nilai yang disimpan. |
| Kolom C | Label yang ditampilkan; jika kosong, nilai pada kolom B digunakan. |
| Kolom D | Nomor urutan; jika kosong, aplikasi memakai urutan `999`. |

Tipe yang dikenali:

- `KATEGORI`
- `PRIORITAS`
- `STATUS`
- `PENYEBAB_KENDALA`

### 6. `TEMPLATE_TASK`

Nama ini digunakan sebagai nama sheet pada template bulk upload yang dibuat aplikasi. Template berisi kolom:

`JUDUL`, `KATEGORI`, `DESKRIPSI`, `JUMLAH_DATA`, `ASSIGN_TO`, `STATUS`, `PRIORITY`, `DUE_DATE`, `ATTACHMENT_LINK`.

## Konfigurasi Google Drive

Folder tujuan lampiran ditentukan oleh `DRIVE_CONFIG.FOLDER_ID` di `Code.gs`. Pastikan akun yang menjalankan/deploy aplikasi mempunyai akses menulis ke folder tersebut. Jika folder tidak dapat dibuka, script mencoba menggunakan root Google Drive akun eksekusi.

## Instalasi di Google Apps Script

1. Buka [Google Apps Script](https://script.google.com/) dan buat project baru.
2. Salin isi `Code.gs` dari repository ke file script bernama `Code.gs`.
3. Tambahkan file HTML bernama `index`, lalu salin isi `index.html` ke file tersebut.
4. Pastikan `CONFIG.SPREADSHEET_ID` menunjuk ke spreadsheet yang benar.
5. Pastikan semua nama sheet dan header sesuai dokumentasi di atas.
6. Pastikan `DRIVE_CONFIG.FOLDER_ID` menunjuk ke folder lampiran yang benar.
7. Simpan project, lalu jalankan `checkSheetStructure()` dari editor untuk memeriksa struktur dasar.
8. Jalankan `runSmokeTest()` bila ingin memeriksa pembacaan pengguna, tugas, statistik, dan kendala. Tinjau hasilnya pada **Execution log**.
9. Berikan izin akses Google Sheets, Google Drive, dan identitas pengguna saat diminta.

## Deploy sebagai Web App

1. Di Google Apps Script, pilih **Deploy > New deployment**.
2. Pilih jenis deployment **Web app**.
3. Tentukan akun pelaksana sesuai kebijakan akses organisasi.
4. Pilih siapa yang boleh mengakses aplikasi.
5. Klik **Deploy**, selesaikan proses otorisasi, lalu buka URL Web App.
6. Setelah kode diperbarui, buat versi deployment baru melalui **Deploy > Manage deployments** agar perubahan tampil pada URL produksi.

Pemilihan akun pelaksana dan cakupan akses menentukan email yang tersedia melalui `Session.getEffectiveUser()`. Uji menggunakan akun pengguna sebenarnya dan pastikan email tersebut tercatat pada sheet `USERS`.

## Cara penggunaan

### Tugas tunggal

1. Buka halaman dashboard.
2. Pilih **Tambah Task**.
3. Isi judul, kategori, deskripsi, jumlah data, penanggung jawab, status, prioritas, tenggat, dan lampiran jika diperlukan.
4. Simpan tugas. Aplikasi membuat `TASK_ID` secara otomatis dengan pola nomor baris, waktu, dan tiga karakter awal nama pembuat.
5. Gunakan aksi pada tabel untuk melihat detail, mengubah, memberi komentar, atau menghapus tugas.

Penghapusan tugas menggunakan mekanisme *soft delete*: baris tetap ada di spreadsheet dan ditandai melalui `IS_DELETED`.

### Bulk upload

1. Unduh template dari menu bulk upload di aplikasi.
2. Isi maksimal 1.000 baris per berkas.
3. Pertahankan nama seluruh header template.
4. Gunakan nama `ASSIGN_TO` yang sama dengan `NAMA_LENGKAP` pada sheet `USERS`.
5. Gunakan nilai kategori, status, dan prioritas yang tersedia pada sheet `CONFIG`.
6. Isi `DUE_DATE` sebagai `dd/mm/yyyy` atau `yyyy-MM-dd`. Kolom template sengaja diformat sebagai teks untuk mencegah perubahan tanggal otomatis oleh Excel/Sheets.
7. Unggah berkas dan periksa ringkasan baris berhasil serta daftar kesalahan. Kesalahan satu baris tidak menghentikan baris lain.

Validasi bulk upload mewajibkan `JUDUL`, `DESKRIPSI`, `JUMLAH_DATA`, `STATUS`, dan `PRIORITY`. `JUMLAH_DATA` harus berupa angka minimal 0.

### Lampiran

Format yang diterima antarmuka meliputi PDF, JPG/JPEG, PNG, DOC/DOCX, XLS/XLSX, dan TXT. File disimpan ke folder Google Drive yang dikonfigurasi, sedangkan URL-nya disimpan pada tugas.

### Follow Up Kendala

Gunakan halaman **Follow Up Kendala** untuk mencatat kendala pada tugas, menentukan penyebab dan status, memperbarui tindak lanjut, menghapus catatan, serta melihat statistik dan filter kendala.

## Pengujian dan pemeliharaan

Fungsi utilitas yang tersedia di editor Apps Script:

| Fungsi | Tujuan |
|---|---|
| `checkSheetStructure()` | Menampilkan header dan jumlah baris pada sheet utama. |
| `runSmokeTest()` | Membaca struktur sheet, pengguna, pengguna aktif, tugas, statistik tugas, dan statistik kendala. |
| `runBulkUploadSmokeTest()` | Menguji bulk upload dengan data contoh, termasuk baris tidak valid. Fungsi ini menulis data uji ke spreadsheet, sehingga gunakan pada salinan/test environment. |

Setelah mengubah data `USERS` atau `CONFIG`, cache dapat bertahan hingga sekitar 300 detik. Tunggu cache berakhir atau jalankan kembali setelah jeda sebelum menyimpulkan bahwa perubahan belum terbaca.

## Pemecahan masalah

### Halaman kosong atau `index` tidak ditemukan

- Pastikan file HTML pada project Apps Script bernama `index`.
- Pastikan isi repository `index.html` disalin ke file tersebut.
- Buat versi deployment baru setelah perubahan.

### Pengguna tidak dikenali

- Cocokkan email akun Google dengan kolom `EMAIL` pada `USERS`.
- Pastikan `USER_ID` tidak kosong.
- Periksa pengaturan **Execute as** dan **Who has access** pada deployment.

### Dropdown kosong

- Pastikan sheet `CONFIG` tersedia.
- Periksa tipe pada kolom A dan nilai pada kolom B.
- Tunggu hingga cache 300 detik berakhir bila konfigurasi baru saja diubah.

### Gagal membaca atau menyimpan data

- Pastikan ID spreadsheet benar dan akun eksekusi memiliki izin edit.
- Pastikan nama sheet dan kapitalisasi header sama persis.
- Jalankan `checkSheetStructure()` dan lihat **Execution log**.

### Lampiran gagal diunggah

- Pastikan ID folder Drive benar.
- Pastikan akun eksekusi mempunyai izin membuat file.
- Periksa izin Drive yang diminta oleh project Apps Script.

### Tanggal tenggat bergeser

- Pada bulk upload, gunakan format `dd/mm/yyyy` atau `yyyy-MM-dd`.
- Jangan mengubah kolom `DUE_DATE` pada template menjadi format tanggal otomatis.
- Pastikan timezone spreadsheet dan project konsisten dengan `Asia/Jakarta`.

## Catatan keamanan

- Jangan publikasikan ID spreadsheet/folder produksi bila repository akan memuat data sensitif.
- Batasi akses Web App sesuai kebutuhan organisasi.
- Atur izin spreadsheet dan folder Drive berdasarkan prinsip akses minimum.
- Validasi kembali siapa yang dapat melihat lampiran yang dibuat di Google Drive.
- Jangan menyimpan kredensial, token, atau data rahasia di source code.

## Checklist deployment

- [ ] `Code.gs` dan `index.html` telah disalin ke project Apps Script.
- [ ] ID spreadsheet dan folder Drive sudah benar.
- [ ] Seluruh sheet dan header sudah tersedia.
- [ ] Pengguna telah terdaftar di `USERS`.
- [ ] Nilai dropdown telah diisi di `CONFIG`.
- [ ] `checkSheetStructure()` tidak melaporkan sheet hilang.
- [ ] `runSmokeTest()` berhasil.
- [ ] Akses tambah/ubah/hapus tugas telah diuji.
- [ ] Bulk upload dan format tanggal telah diuji.
- [ ] Komentar, kendala, dan lampiran telah diuji.
- [ ] Deployment Web App terbaru sudah aktif.

## Lisensi

Belum ada file lisensi pada repository ini. Tambahkan lisensi sebelum mendistribusikan atau menggunakan kode di luar kebutuhan internal.
