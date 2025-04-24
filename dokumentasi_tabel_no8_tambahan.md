# Dokumentasi Script `tabel_no8_tambahan.py`

## Tujuan Program
Script `tabel_no8_tambahan.py` dirancang untuk membandingkan hasil analisis varians (ANOVA) yang dihitung oleh program dengan tabel referensi yang telah disediakan. Program ini melakukan perhitungan ANOVA untuk eksperimen faktorial dengan tiga faktor (Komposisi, Kompaksi, Cavity) pada berbagai frekuensi, lalu membandingkan hasilnya dengan tabel referensi untuk memvalidasi keakuratan perhitungan.

## Library yang Digunakan
- **pandas**: Untuk manipulasi dan analisis data dalam bentuk DataFrame
- **numpy**: Untuk perhitungan numerik dan statistik dasar
- **scipy.stats**: Untuk menghitung nilai kritis F dalam analisis ANOVA

## Struktur Program

### Fungsi Utama
Program ini memiliki satu fungsi utama `compare_anova_results()` yang mengkoordinasikan seluruh proses:

1. **Membaca Data Input**
   - Membaca tabel referensi dari `input/semua_tabel.xlsx` (sheet 'tabel_4.6')
   - Membaca data eksperimen dari `input/tabel_koef_Serap_bunyi.xlsx` (sheet 'CRD')

2. **Pengkodean Faktor**
   - Mengubah nilai faktor eksperimen ke dalam bentuk kode (-1, 0, 1):
     - Komposisi: "50 : 50" = -1, "70 : 30" = 0, "90 : 10" = 1
     - Kompaksi: "3 : 4" = -1, "4 : 4" = 0, "5 : 4" = 1
     - Cavity: "15 mm" = -1, "20 mm" = 0, "25 mm" = 1

3. **Perhitungan ANOVA per Frekuensi**
   - Untuk setiap frekuensi unik dalam data eksperimen:
     - Memfilter data untuk frekuensi tersebut
     - Melakukan perhitungan ANOVA melalui fungsi `calculate_anova_for_frequency`

4. **Fungsi `calculate_anova_for_frequency`**
   - Menghitung komponen-komponen ANOVA:
     - Sum of Squares (SS): Total, Treatment, Error, dan untuk setiap faktor dan interaksinya
     - Degrees of Freedom (df): Untuk setiap faktor dan interaksinya
     - Mean Squares (MS): SS dibagi dengan df masing-masing
     - F-values (F): MS untuk faktor dibagi dengan MS error
     - F-critical values: Nilai ambang batas signifikansi untuk setiap faktor
     - Penentuan signifikansi: Faktor signifikan jika F > F-critical

5. **Perbandingan dengan Referensi**
   - Membandingkan hasil perhitungan ANOVA untuk frekuensi "Rata-rata" dengan tabel referensi
   - Memeriksa kecocokan untuk setiap parameter (df, SS, MS, F-values, dan signifikansi)

6. **Penyimpanan Hasil**
   - Menyimpan hasil ANOVA untuk semua frekuensi ke dalam file Excel
   - Menyimpan tabel perbandingan dengan referensi
   - Menyimpan tabel referensi dan data eksperimen lengkap
   - Output disimpan di `output/nomor8_tambahan.xlsx`

7. **Pelaporan Hasil**
   - Menampilkan ringkasan perbandingan dengan tabel referensi
   - Menampilkan jumlah frekuensi yang dianalisis dan total data yang diolah

## Detail Perhitungan ANOVA

### Komponen Perhitungan ANOVA
1. **Sum of Squares (SS)**
   - **SS Total**: Mengukur total variasi dalam data
     ```python
     ss_total = sum((data['Nilai Respon'] - grand_mean) ** 2)
     ```
   
   - **SS Treatment**: Mengukur variasi antar kombinasi perlakuan
     ```python
     ss_treatment = sum((treatment_means - grand_mean) ** 2) * n_rep
     ```
   
   - **SS Main Effects**: Mengukur variasi antar level satu faktor
     ```python
     def calculate_main_effect_ss(factor):
         factor_means = data.groupby(factor)['Nilai Respon'].mean()
         return sum((factor_means - grand_mean) ** 2) * (n_total / len(factor_means))
     ```
   
   - **SS Interaction**: Mengukur variasi karena interaksi antar faktor
     ```python
     def calculate_interaction_ss(factor1, factor2):
         interaction_means = data.groupby([factor1, factor2])['Nilai Respon'].mean()
         n_combinations = len(interaction_means)
         return sum((interaction_means - grand_mean) ** 2) * (n_total / n_combinations) - \
                calculate_main_effect_ss(factor1) - calculate_main_effect_ss(factor2)
     ```
   
   - **SS Error**: Mengukur variasi yang tidak dapat dijelaskan oleh model
     ```python
     ss_error = ss_total - ss_treatment
     ```

2. **Degrees of Freedom (df)**
   - Setiap faktor: jumlah level - 1 = 2 (karena 3 level)
   - Interaksi: df faktor 1 × df faktor 2
   - Error: jumlah total observasi - jumlah perlakuan
   - Total: jumlah total observasi - 1

3. **Mean Squares (MS)**
   - MS = SS / df untuk masing-masing komponen

4. **F-values**
   - F = MS faktor / MS error
   - Signifikansi ditentukan dengan membandingkan F-value dengan F-critical

## Format Output
Output dari program ini adalah file Excel dengan beberapa sheet:
1. **ANOVA_{Frekuensi}**: Hasil ANOVA untuk setiap frekuensi
2. **Perbandingan**: Tabel perbandingan hasil ANOVA dengan tabel referensi
3. **Tabel 4.6 Reference**: Data referensi dari literatur
4. **Data_Lengkap**: Data eksperimen yang digunakan dalam perhitungan

## Penggunaan Program
Program dapat dijalankan dengan perintah:
```
python tabel_no8_tambahan.py
```

Program mengasumsikan bahwa:
1. Tabel referensi tersedia di `input/semua_tabel.xlsx` (sheet 'tabel_4.6')
2. Data eksperimen tersedia di `input/tabel_koef_Serap_bunyi.xlsx` (sheet 'CRD')
3. Direktori `output` telah tersedia untuk menyimpan hasil

## Penanganan Error
Program dilengkapi dengan mekanisme penanganan kesalahan yang akan menampilkan pesan error dan detail traceback jika terjadi masalah saat eksekusi. 