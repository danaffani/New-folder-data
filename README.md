# Analisis Data Eksperimen

Repository ini berisi kode Python untuk menganalisis data eksperimen dari skripsi tentang material penyerap suara.

[🔍 Lihat Hasil Analisis Lengkap](analysis_answers.md)

## Dependencies

Pastikan Anda telah menginstall Python 3.7+ dan package berikut:
```
pandas>=1.3.0
numpy>=1.20.0
scipy>=1.7.0
statsmodels>=0.13.0
matplotlib>=3.4.0
seaborn>=0.11.0
```

## Persiapan Environment

1. Clone repository ini:
```bash
git clone <repository_url>
cd <repository_name>
```

2. Buat dan aktifkan virtual environment:
```bash
# Windows
python -m venv .venv
.venv\Scripts\activate

# Linux/Mac
python3 -m venv .venv
source .venv/bin/activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Verifikasi instalasi:
```bash
python check_dependencies.py
```
Jika semua dependencies terinstall dengan benar, Anda akan melihat pesan konfirmasi.

## Struktur File

```
.
├── input/                      # Data input
│   ├── semua_tabel.xlsx       # File data utama
│   └── ...
├── output/                    # Hasil analisis
├── tabel_no3_tambahan.py     # Generate design matrix
├── tabel_no8_tambahan.py     # Analisis ANOVA
├── analysis_answers.md       # Penjelasan analisis
├── requirements.txt          # Daftar dependencies
└── README.md                # Dokumentasi
```

## Penggunaan

1. **Persiapkan Data Input**:
   - Pastikan file `semua_tabel.xlsx` ada di folder `input/`
   - File harus memiliki sheet 'tabel_4.6' untuk referensi ANOVA

2. **Generate Design Matrix**:
   ```bash
   python tabel_no3_tambahan.py
   ```
   Menghasilkan file `output/nomor3_tambahan.xlsx` dengan design matrix RBD dan CRD.

3. **Analisis ANOVA**:
   ```bash
   python tabel_no8_tambahan.py
   ```
   Menghasilkan file `output/nomor8_tambahan.xlsx` dengan hasil ANOVA dan perbandingan.

4. **Lihat Hasil**:
   - Buka file Excel di folder `output/`
   - Baca penjelasan analisis di `analysis_answers.md`

## Troubleshooting

Jika mengalami error:

1. **ModuleNotFoundError**:
   - Pastikan virtual environment aktif
   - Jalankan `pip install -r requirements.txt` kembali
   - Verifikasi dengan `python check_dependencies.py`

2. **FileNotFoundError**:
   - Periksa struktur folder dan file input
   - Pastikan nama file dan path sudah benar

3. **Error lainnya**:
   - Periksa log error yang ditampilkan
   - Pastikan format data input sesuai
   - Jika masih bermasalah, buat issue baru di repository

## Lisensi

MIT License - lihat file LICENSE untuk detail. 