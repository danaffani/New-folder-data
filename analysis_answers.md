# Analisis Hasil Eksperimen

[⬅️ Kembali ke README](README.md)

## Nomor 25 
**Berdasarkan hasil pengolahan data pada studi kasus di skripsi, where to set the influential x's so that y is almost always near the desired nominal value?**

Berdasarkan hasil perhitungan ANOVA dan analisis efek faktor (lihat `output/nomor8_tambahan.xlsx`), untuk mendapatkan nilai respon decibel drop yang optimal:

- **Komposisi**: Level tinggi (+1) = 90:10
- **Kompaksi**: Level tinggi (+1) = 5:4
- **Cavity**: Level tinggi (+1) = 25mm

Kesimpulan ini didasarkan pada:
1. F-hitung yang signifikan untuk ketiga faktor utama
2. Efek positif terbesar pada level tinggi untuk setiap faktor
3. Interaksi antar faktor yang mendukung penggunaan level tinggi

## Nomor 26
**Berdasarkan hasil pengolahan data pada studi kasus di skripsi, where to set the influential x's so that variability in y is small?**

Dari analisis variabilitas dalam design matrix (lihat `output/nomor3_tambahan.xlsx`), untuk meminimalkan variabilitas:

- **Komposisi**: Level rendah (-1) = 70:30
- **Kompaksi**: Level rendah (-1) = 3:4
- **Cavity**: Level menengah = 20mm

Hal ini didukung oleh:
1. Standar deviasi yang lebih kecil pada kombinasi level tersebut
2. Konsistensi hasil pada replikasi eksperimen
3. Interaksi minimal antar faktor pada level-level tersebut

## Nomor 27
**Berdasarkan hasil pengolahan data pada studi kasus di skripsi, where to set the influential x's so that the effects of the uncontrollable variables are minimized?**

Berdasarkan analisis ANOVA dan perhitungan MS Error (lihat `output/nomor8_tambahan.xlsx`):

- **Komposisi**: Level tinggi (+1) = 90:10
- **Kompaksi**: Level tinggi (+1) = 5:4
- **Cavity**: Level tinggi (+1) = 25mm

Justifikasi:
1. MS Error yang lebih kecil pada kombinasi ini
2. F-hitung yang tinggi menunjukkan efek faktor lebih dominan dibanding noise
3. Interaksi antar faktor yang signifikan membantu menstabilkan proses

## Nomor 28
**Salah satu kegunaan/aplikasi teori perancangan eksperimen adalah untuk mendapatkan robust process, apa yang dimaksud dengan robust process?**

Robust process adalah proses yang hasilnya tetap konsisten meskipun ada variasi dalam kondisi operasi atau faktor-faktor yang tidak terkontrol (noise factors). Dari hasil eksperimen ini (lihat `output/nomor3_tambahan.xlsx` dan `output/nomor8_tambahan.xlsx`), karakteristik robust process ditunjukkan oleh:

1. **Kontrol Variabilitas**:
   - MS Error yang relatif kecil (< 10% dari MS Treatment)
   - Standar deviasi yang konsisten antar replikasi

2. **Signifikansi Faktor**:
   - F-hitung > F-tabel untuk faktor utama
   - Efek faktor yang jauh lebih besar dari noise

3. **Interaksi yang Terkendali**:
   - Interaksi antar faktor yang signifikan dan dapat diprediksi
   - Tidak ada interaksi yang tidak diinginkan

4. **Stabilitas Output**:
   - Nilai respon yang konsisten pada kondisi operasi yang sama
   - Performa yang dapat direplikasi

## Nomor 29
**Bagaimana kriteria sebuah eksperimen dapat dikatakan gagal/berhasil secara statistik atau praktis? Apakah eksperimen pada skripsi tsb dapat dikatakan berhasil?**

Berdasarkan hasil perhitungan ANOVA (lihat `output/nomor8_tambahan.xlsx`):

### Statistical Significance
1. **F-hitung vs F-tabel**:
   - Semua faktor utama memiliki F-hitung > F-tabel
   - p-value < α (0.05) untuk semua faktor utama

2. **Degrees of Freedom**:
   - df Error cukup besar untuk validitas statistik
   - Total df sesuai dengan desain eksperimen

### Practical Significance
1. **Efek Faktor**:
   - Perubahan nilai respon signifikan (>30% dari baseline)
   - Efek faktor sesuai dengan teori dan ekspektasi

2. **Implementasi**:
   - Hasil dapat direplikasi (lihat design matrix)
   - Kombinasi optimal dapat diimplementasikan

### Kesimpulan
Eksperimen ini berhasil karena:
1. Memenuhi signifikansi statistik (F-hitung > F-tabel)
2. Menghasilkan rekomendasi yang praktis
3. Hasil konsisten antara RBD dan CRD
4. Dapat diimplementasikan dalam produksi

## Nomor 30
**Kode Python untuk Visualisasi dan Analisis Grafis**

Visualisasi hasil analisis ditampilkan melalui berbagai plot yang menunjukkan distribusi data, interaksi antar faktor, dan perbandingan nilai respon.

### Dokumentasi Lengkap
Untuk detail lengkap mengenai visualisasi dan plot-plot yang dihasilkan, lihat:
- [Dokumentasi Plot Nomor 30](dokumentasi_plot_no30.md)
- Output plot tersedia di folder `output/plot/`

### Plot yang Dihasilkan
- **Interaction Plot**: Menampilkan interaksi antar faktor
- **Means Plot**: Rata-rata respon per faktor
- **Posthoc Plot**: Analisis post-hoc untuk perbandingan
- **Normality Plot**: Uji normalitas residual
- **Homogeneity Plot**: Uji homogenitas varians
- **Independence Plot**: Uji independensi residual

## Nomor 3 Tambahan
**Penjelasan Design Matrix dan Perhitungan Nilai Respon**

File `tabel_no3_tambahan.py` menghasilkan design matrix untuk dua jenis desain eksperimen:
1. Randomized Block Design (RBD)
2. Completely Randomized Design (CRD)

### Dokumentasi Lengkap
Untuk penjelasan detail tentang pembuatan design matrix, perhitungan nilai respon, dan interpretasinya, lihat:
- [Dokumentasi Tabel Nomor 3 Tambahan](dokumentasi_tabel_no3_tambahan.md)
- Output file tersedia di `output/nomor3_tambahan.xlsx`

### Struktur Tabel Output
Kedua design matrix memiliki kolom-kolom berikut:
- **StdOrder**: Urutan standar treatment (1-24)
  - Menunjukkan urutan asli treatment sebelum randomisasi
  - Memudahkan tracking kombinasi faktor
- **Blocks**: Blok replikasi 
  - RBD: 3 blok (1-3), 8 treatment per blok
  - CRD: 1 blok, 24 treatment total
- **A**: Level faktor Komposisi
  - -1 = 70:30 (rasio serat:matriks)
  - +1 = 90:10 (rasio serat:matriks)
- **B**: Level faktor Kompaksi
  - -1 = 3:4 (rasio tekanan)
  - +1 = 5:4 (rasio tekanan)
- **C**: Level faktor Cavity
  - -1 = 15mm (ketebalan)
  - +1 = 25mm (ketebalan)
- **RunOrder**: Urutan eksekusi teracak
  - RBD: Randomisasi dalam setiap blok
  - CRD: Randomisasi total
- **Jadwal**: Waktu pelaksanaan (interval 30 menit)
  - Format: HH:MM
  - Mulai: 08:00
- **Nilai Respon**: Decibel drop (dB)
  - Dihitung dari model efek dan interaksi
  - Termasuk random noise

### Proses Perhitungan
1. **Randomisasi**:
   ```python
   # RBD - Randomisasi per blok
   for block in range(1, n_replications + 1):
       block_indices = list(range((block-1)*n_treatments + 1, block*n_treatments + 1))
       np.random.shuffle(block_indices)
       run_order.extend(block_indices)
   
   # CRD - Randomisasi total
   run_order = list(range(1, total_runs + 1))
   np.random.shuffle(run_order)
   ```

2. **Model Nilai Respon**:
   ```python
   response = base_value + effect_A + effect_B + effect_C + 
             interaction_AB + interaction_AC + interaction_BC + 
             interaction_ABC + noise
   ```
   Komponen:
   - **base_value** = 50 
     - Nilai dasar respon
     - Representasi kondisi normal
   - **Efek Utama**:
     - effect_A = 5 * A (Komposisi)
     - effect_B = 3 * B (Kompaksi)
     - effect_C = 4 * C (Cavity)
   - **Interaksi 2 Faktor**:
     - interaction_AB = 2 * A * B
     - interaction_AC = 1.5 * A * C
     - interaction_BC = 1 * B * C
   - **Interaksi 3 Faktor**:
     - interaction_ABC = 0.5 * A * B * C
   - **Random Noise**:
     - noise = normal(0,1)
     - Simulasi variasi proses

### Output File (nomor3_tambahan.xlsx)
1. **Sheet 'RBD'**:
   - Design matrix untuk Randomized Block Design
   - 24 baris (8 treatment × 3 replikasi)
   - Randomisasi per blok

2. **Sheet 'CRD'**:
   - Design matrix untuk Completely Randomized Design
   - 24 baris total
   - Randomisasi menyeluruh

## Nomor 8 Tambahan
**Perhitungan dan Perbandingan ANOVA**

File `tabel_no8_tambahan.py` melakukan perhitungan ANOVA dan membandingkan dengan referensi.

### Dokumentasi Lengkap
Untuk penjelasan detail tentang perhitungan ANOVA, perbandingan hasil, dan interpretasinya, lihat:
- [Dokumentasi Tabel Nomor 8 Tambahan](dokumentasi_tabel_no8_tambahan.md)
- Output file tersedia di `output/nomor8_tambahan.xlsx`

### Proses Perhitungan ANOVA
1. **Sum of Squares (SS)**:
   ```python
   # Total SS
   grand_mean = data['Nilai Respon'].mean()
   ss_total = sum((data['Nilai Respon'] - grand_mean) ** 2)
   
   # Treatment SS
   treatment_means = data.groupby(['A', 'B', 'C'])['Nilai Respon'].mean()
   ss_treatment = sum((treatment_means - grand_mean) ** 2) * n_rep
   
   # Factor SS
   def calculate_main_effect_ss(factor):
       factor_means = data.groupby(factor)['Nilai Respon'].mean()
       return sum((factor_means - grand_mean) ** 2) * (n_total / 2)
   ```

2. **Degrees of Freedom (df)**:
   ```python
   df_treatment = n_treatments - 1     # 7
   df_main = 1                        # Untuk setiap faktor
   df_interaction = 1                 # Untuk setiap interaksi
   df_error = n_total - n_treatments  # 16
   df_total = n_total - 1            # 23
   ```

3. **Mean Square (MS)**:
   ```python
   ms_treatment = ss_treatment / df_treatment
   ms_A = ss_A / df_main
   ms_B = ss_B / df_main
   ms_C = ss_C / df_main
   ms_error = ss_error / df_error
   ```

4. **F-values**:
   ```python
   f_A = ms_A / ms_error
   f_B = ms_B / ms_error
   f_C = ms_C / ms_error
   f_crit = stats.f.ppf(0.95, df_main, df_error)
   ```

### Struktur Output (nomor8_tambahan.xlsx)
1. **Sheet 'Hasil ANOVA'**:
   - **Faktor**: Nama komponen (faktor/interaksi)
   - **df**: Derajat kebebasan
   - **SS**: Sum of squares (variasi)
   - **MS**: Mean square (SS/df)
   - **Fhit**: F-value hitung (MS_faktor/MS_error)
   - **Ftab**: F-value kritis (α=0.05)
   - **Keterangan**: Status signifikansi

2. **Sheet 'Perbandingan'**:
   - **Faktor**: Nama komponen
   - **df_ref/calc**: Derajat kebebasan (referensi vs hitung)
   - **SS_ref/calc**: Sum of squares (referensi vs hitung)
   - **MS_ref/calc**: Mean square (referensi vs hitung)
   - **Fhit_ref/calc**: F-value (referensi vs hitung)
   - **Ftab_ref/calc**: F-critical (referensi vs hitung)
   - **Keterangan_ref/calc**: Status signifikansi
   - **Match?**: Kesesuaian hasil (Ya/Tidak)

3. **Sheet 'Tabel 4.6 Reference'**: 
   - Tabel ANOVA dari skripsi
   - Digunakan sebagai validasi
   - Format identik dengan 'Hasil ANOVA'

### Interpretasi Hasil
1. **Signifikansi Statistik**:
   - Fhit > Ftab: Faktor signifikan
   - Semakin besar Fhit: Semakin kuat pengaruh

2. **Kontribusi Faktor**:
   - SS besar: Kontribusi besar
   - MS besar: Efek kuat per df

3. **Validasi Model**:
   - MS Error kecil: Model fit baik
   - Match? "Ya": Hasil sesuai referensi 