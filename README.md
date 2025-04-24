# Analisis Data Eksperimen Material Penyerap Suara

Repository ini berisi kode Python untuk menganalisis data eksperimen dari skripsi tentang material penyerap suara. Analisis meliputi pembuatan design matrix, analisis ANOVA, dan visualisasi data.

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

## Struktur Proyek

```
.
├── input/                     # Data input
│   ├── semua_tabel.xlsx       # File data utama
│   ├── tabel_koef_Serap_bunyi.xlsx
│   ├── design_matrix.xlsx
│   ├── nuisance_factors.xlsx
│   └── hypothesis_table.xlsx
├── output/                    # Hasil analisis
│   ├── nomor3_tambahan.xlsx   # Design matrix RBD dan CRD
│   ├── nomor8_tambahan.xlsx   # Hasil ANOVA
│   └── plot/                  # Visualisasi dan plot
│       ├── interaction_plot.png
│       ├── means_plot.png
│       ├── posthoc_plot.png
│       └── ...
├── raw_input/                 # Data mentah
│   ├── generate_table_input.py
│   ├── create_table_lampiran1.py
│   ├── create_table_lampiran1_revised.py
│   └── tabel_*.txt            # File teks data sumber
├── tabel_no3_tambahan.py      # Generate design matrix
├── tabel_no8_tambahan.py      # Analisis ANOVA
├── plot_no30.py               # Script visualisasi data
├── create_table_lampiran1.py  # Generate tabel lampiran
├── dokumentasi_*.md           # Dokumentasi proses analisis
├── analysis_answers.md        # Penjelasan analisis
└── README.md                  # Dokumentasi
```

## Penggunaan

1. **Persiapkan Data Input**:
   - Pastikan folder `input/` berisi file Excel yang diperlukan
   - File `semua_tabel.xlsx` harus memiliki sheet yang sesuai untuk analisis

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

4. **Membuat Visualisasi**:
   ```bash
   python plot_no30.py
   ```
   Menghasilkan visualisasi di folder `output/plot/`.

5. **Membuat Tabel Lampiran**:
   ```bash
   python create_table_lampiran1.py
   ```

## Dokumentasi

Dokumentasi detail untuk masing-masing proses analisis tersedia dalam file berikut:
- `dokumentasi_tabel_no3_tambahan.md` - Proses pembuatan design matrix
- `dokumentasi_tabel_no8_tambahan.md` - Analisis ANOVA dan interpretasi
- `dokumentasi_plot_no30.md` - Visualisasi data dan hasilnya
- `dokumentasi_create_table_lampiran1.md` - Pembuatan tabel lampiran

## Troubleshooting

Jika mengalami error:

1. **ModuleNotFoundError**:
   - Pastikan package yang diperlukan telah diinstall
   - Jalankan `pip install <nama_package>` untuk menginstall package yang diperlukan

2. **FileNotFoundError**:
   - Periksa struktur folder dan file input
   - Pastikan nama file dan path sudah benar

3. **Error lainnya**:
   - Periksa log error yang ditampilkan
   - Pastikan format data input sesuai

## Lisensi

MIT License - lihat file LICENSE untuk detail. 