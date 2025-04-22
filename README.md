# Proyek Analisis Eksperimen

## Struktur Proyek

Proyek ini terdiri dari 3 folder utama:

1. `raw_input` - Berisi file tabel_*.txt (data teks mentah) dan script untuk mengkonversi ke Excel
2. `input` - Berisi file Excel yang dibutuhkan untuk analisis
3. `output` - Berisi hasil output dari semua analisis

## Alur Penggunaan

1. Pertama, jalankan `project_setup.py` untuk membuat struktur folder dan memindahkan file teks ke folder raw_input:
   ```
   python project_setup.py
   ```

2. Konversi semua file teks ke format Excel dengan menjalankan:
   ```
   python convert_txt_to_excel.py
   ```
   Script ini akan membaca semua file tabel_*.txt dari folder raw_input dan menyimpannya sebagai Excel di folder input.

3. Buat design matrix dan tabel hipotesis dengan menjalankan:
   ```
   python design_matrix_generator.py
   ```
   File output akan disimpan di folder output.

4. Jalankan analisis ANOVA dan visualisasi dengan:
   ```
   python anova_analysis.py
   ```
   File output (hasil ANOVA, post-hoc tests) akan disimpan di folder output.

## Deskripsi File Utama

* `project_setup.py` - Script untuk persiapan struktur proyek
* `convert_txt_to_excel.py` - Mengkonversi file teks (.txt) ke Excel (.xlsx)
* `design_matrix_generator.py` - Membuat design matrix untuk RBD dan CRD beserta tabel hipotesis
* `anova_analysis.py` - Melakukan analisis ANOVA, post-hoc tests, uji asumsi, dan visualisasi

## Persyaratan

* Python 3.6 atau lebih tinggi
* Library yang dibutuhkan:
  - pandas
  - numpy
  - matplotlib
  - seaborn
  - openpyxl
  - statsmodels
  - scipy

Anda dapat menginstal semua library yang dibutuhkan dengan:
```
pip install pandas numpy matplotlib seaborn openpyxl statsmodels scipy
```

## Hasil Analisis

Hasil analisis akan tersimpan di folder output dengan struktur:
* `output/results/` - File CSV hasil analisis ANOVA dan post-hoc tests
* `output/plots/` - Visualisasi dan grafik
* `output/analysis_summary.xlsx` - Ringkasan hasil analisis
* `output/design_matrices.xlsx` - Design matrix untuk RBD dan CRD
* `output/hypothesis_table.xlsx` - Tabel hipotesis penelitian

## Pertanyaan Penelitian

Jawaban untuk pertanyaan penelitian nomor 25-30 dapat dilihat di file `analysis_answers.md`. 