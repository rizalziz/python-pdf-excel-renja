# PDF to Excel Converter - SIPD-RI

Aplikasi desktop berbasis Python untuk mengubah file PDF yang berisi tabel menjadi file Excel (.xlsx) secara otomatis. Aplikasi ini dibuat dengan antarmuka grafis menggunakan Tkinter dan memanfaatkan `pdfplumber`, `pandas`, serta `openpyxl` untuk membaca isi PDF, membersihkan data, dan mengekspor ke spreadsheet Excel.

## Fitur Utama

- Memindai file PDF yang ada di folder kerja saat ini
- Menampilkan daftar PDF yang tersedia dalam combo box
- Memilih satu file PDF untuk dikonversi
- Menghapus footer berulang seperti `SIPD-RI***` secara opsional
- Mengatur lebar kolom Excel secara otomatis
- Membaca tabel dari halaman PDF dan menggabungkannya ke dalam satu sheet Excel
- Menghapus header yang berulang pada halaman berikutnya
- Membersihkan data numerik seperti ribuan dan tanda koma
- Menghapus baris nomor dan baris kosong yang tidak relevan
- Mendeteksi struktur data induk-anak dan menambahkan rumus penjumlahan otomatis untuk kolom tertentu
- Menyimpan file hasil konversi dengan nama yang sama seperti file PDF asal

## Teknologi yang Digunakan

- Python (FastAPI) untuk backend API
- React + TypeScript (Vite) untuk antarmuka pengguna
- pdfplumber untuk mengekstrak tabel dari PDF
- pandas untuk pemrosesan data tabular
- openpyxl untuk membuat dan memformat file Excel
- Cloudflare Turnstile untuk verifikasi captcha
- Docker & Docker Compose untuk deployment
- nginx untuk menyajikan frontend dan mem-proxy `/api` ke backend

## Persyaratan Sistem

- Docker dan Docker Compose
- File `.env` yang berisi:
  - `TURNSTILE_SITE_KEY`
  - `TURNSTILE_SECRET_KEY`

## Instalasi & Cara Menjalankan

1. Pastikan file `.env` sudah dibuat di folder project dengan konfigurasi Turnstile.
2. Bangun dan jalankan semua service dengan Docker Compose:

```bash
docker compose up --build
```

3. Buka frontend di browser:

```text
http://localhost:3000
```

4. API backend dapat diakses di port `8501` (misal `http://localhost:8501/api/config`).

Cara menggunakan aplikasi:

1. Selesaikan verifikasi Captcha Turnstile.
2. Pilih file PDF yang ingin dikonversi.
3. Atur opsi konversi sesuai kebutuhan:
   - hapus footer SIPD-RI
   - auto lebar kolom Excel
4. Klik tombol `Konversi ke Excel`.
5. Klik `Download File Excel` untuk menyimpan hasilnya.

## Alur Kerja Aplikasi

Aplikasi melakukan proses berikut:

1. Mencari file `.pdf` di folder kerja saat ini.
2. Membaca halaman PDF satu per satu menggunakan `pdfplumber`.
3. Mengekstrak tabel yang ada di halaman tersebut.
4. Membersihkan baris kosong, footer, dan header yang berulang.
5. Membersihkan kolom data numerik agar format Excel benar.
6. Menyusun data menjadi DataFrame pandas.
7. Menulis hasil ke workbook Excel menggunakan `openpyxl`.
8. Menerapkan format kolom dan border sesuai kebutuhan.
9. Menambahkan rumus `SUM` untuk data yang bersifat hierarki/anak.

## Struktur Data Hasil Excel

Hasil konversi akan dibuat dalam sheet bernama `Data PDF`, dan biasanya berisi:

- header tabel dari PDF
- baris data utama
- data numerik yang sudah dibersihkan
- rumus total pada data yang memiliki hubungan induk-anak

## Catatan Penting

- Aplikasi hanya men-scan file PDF di folder kerja saat ini (`os.getcwd()`).
- Output akan disimpan di lokasi yang sama dengan file PDF.
- Kinerja terbaik didapatkan jika PDF memiliki tabel yang terstruktur dengan baik.
- Jika PDF tidak memiliki tabel yang dapat terbaca, aplikasi akan menampilkan pesan bahwa tidak ada tabel yang ditemukan.
- File log proses akan dibuat dengan nama `converter.log` di folder project.

## Contoh Penggunaan

Misalnya Anda memiliki file berikut di folder project:

```text
laporan-sipd.pdf
```

Setelah menjalankan aplikasi dan menekan tombol konversi, maka akan terbentuk file:

```text
laporan-sipd.xlsx
```

## Lisensi

Project ini dibuat untuk kebutuhan konversi data PDF ke Excel dan dapat disesuaikan sesuai kebutuhan internal penggunaan.

## Penutup

Aplikasi ini sangat cocok untuk kebutuhan konversi tabel PDF yang memiliki format relatif konsisten, khususnya untuk dokumen SIPD-RI yang umum dipakai dalam data perencanaan dan pelaporan.
