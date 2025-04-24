# Dokumentasi Script `create_table_lampiran1.py`

## Tujuan Program
Script `create_table_lampiran1.py` dibuat untuk mengolah data hasil eksperimen faktor-faktor yang mempengaruhi koefisien serap bunyi. Program menghasilkan tabel-tabel yang dibutuhkan untuk lampiran, dengan mempertimbangkan 5 spesimen (blocks) untuk setiap kombinasi faktor, dan menyusun data sesuai dengan format Randomized Block Design (RBD) dan Completely Randomized Design (CRD).

## Library yang Digunakan
- **pandas**: Untuk manipulasi dan analisis data tabular
- **numpy**: Untuk operasi matematika dan pengacakan data
- **os**: Untuk operasi sistem seperti pembuatan direktori

## Struktur Program

### Fungsi Utama
Program memiliki tiga fungsi utama:

#### `buat_tabel_lampiran1()`
- Fungsi utama yang mengkoordinasikan seluruh proses
- Membuat direktori `output` jika belum ada
- Membaca data dari file Excel `input/tabel_koef_Serap_bunyi_revised.xlsx`
- Memanggil fungsi-fungsi lain untuk memproses data
- Menangani error jika terjadi masalah saat membaca file

#### `buat_tabel_koefisien_serap_bunyi_per_frekuensi(df_rbd)`
- Menerima DataFrame dari data RBD
- Memproses data untuk setiap frekuensi yang ada
- Mengelompokkan data berdasarkan kombinasi faktor (Komposisi, Kompaksi, Cavity)
- Mengambil nilai spesimen dari 5 blocks untuk setiap kombinasi
- Menghitung rata-rata dan standar deviasi dari 5 nilai spesimen
- Menghasilkan tabel lengkap yang menampilkan:
  - Frekuensi
  - Komposisi
  - Kompaksi
  - Cavity
  - Nilai individual untuk Spesimen 1-5
  - Rata-rata nilai
  - Standar deviasi
- Menyimpan hasil ke file `output/tabel_koefisien_serap_bunyi_per_frekuensi.xlsx`

#### `buat_tabel_rbd_crd(df_rbd)`
- Menerima DataFrame dari data RBD
- Menyalin data untuk format RBD (Randomized Block Design)
- Membuat data format CRD (Completely Randomized Design) dengan mengacak RunOrder
- Menyimpan kedua tabel ke file `output/tabel_rbd_crd.xlsx` dalam dua sheet berbeda

## Struktur Data
Data yang diproses memiliki struktur sebagai berikut:
- **StdOrder**: Urutan standar eksperimen
- **RunOrder**: Urutan pelaksanaan eksperimen
- **Blocks**: Representasi 5 spesimen yang digunakan (1-5)
- **Komposisi**: Faktor level komposisi material
- **Kompaksi**: Faktor level kompaksi material
- **Cavity**: Faktor level cavity
- **Frekuensi**: Frekuensi pengukuran
- **Nilai Spesimen**: Nilai koefisien serap bunyi yang diukur
- **Rata-rata Koefisien**: Rata-rata nilai koefisien

## Desain Eksperimen
Program ini mengasumsikan eksperimen dengan desain faktorial penuh dengan:
- 5 spesimen (blocks) untuk setiap kombinasi perlakuan
- 3 faktor (Komposisi, Kompaksi, Cavity) dengan level bervariasi
- Pengukuran pada beberapa frekuensi yang berbeda

## Penggunaan Program
1. Pastikan file data input tersedia di `input/tabel_koef_Serap_bunyi_revised.xlsx`
2. Jalankan program dengan perintah `python create_table_lampiran1.py`
3. Program akan menghasilkan dua file output:
   - `output/tabel_koefisien_serap_bunyi_per_frekuensi.xlsx`: Tabel yang menampilkan nilai 5 spesimen untuk setiap kombinasi faktor pada setiap frekuensi
   - `output/tabel_rbd_crd.xlsx`: Tabel data dalam format RBD dan CRD

## Catatan Penting
- Blocks dalam data mewakili 5 spesimen berbeda, bukan replikasi
- Program mengasumsikan bahwa data sudah tersusun dengan benar dalam sheet RBD di file input
- Jika ada nilai yang hilang untuk spesimen tertentu, program akan menyimpannya sebagai NaN 