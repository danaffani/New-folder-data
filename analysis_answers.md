# Jawaban Soal Thesis

## Nomor 25 
**Berdasarkan hasil pengolahan data pada studi kasus di skripsi, where to set the influential x's so that y is almost always near the desired nominal value?**

Berdasarkan hasil analisis, untuk mendapatkan nilai respon decibel drop yang optimal (nilai tertinggi):

- **Komposisi**: Gunakan komposisi 90:10 (level tinggi)
- **Kompaksi**: Gunakan kompaksi 5:4 (level tinggi)
- **Cavity**: Gunakan cavity 25mm (level tinggi)

Kombinasi ini akan menghasilkan nilai decibel drop yang paling tinggi dan konsisten, yang berarti performa penyerapan suara yang lebih baik.

## Nomor 26
**Berdasarkan hasil pengolahan data pada studi kasus di skripsi, where to set the influential x's so that variability in y is small?**

Untuk meminimalkan variabilitas dalam nilai respon:

- **Komposisi**: 70:30 (level menengah)
- **Kompaksi**: 4:4 (level menengah)
- **Cavity**: 20mm (level menengah)

Kombinasi ini menunjukkan variansi yang paling kecil di antara semua replikasi, yang menandakan proses yang lebih stabil dan hasil yang lebih konsisten.

## Nomor 27
**Berdasarkan hasil pengolahan data pada studi kasus di skripsi, where to set the influential x's so that the effects of the uncontrollable variables are minimized?**

Untuk meminimalkan efek dari variabel yang tidak terkontrol:

- **Komposisi**: 90:10 (level tinggi)
- **Kompaksi**: 5:4 (level tinggi)
- **Cavity**: 25mm (level tinggi)

Pada level-level ini, pengaruh dari noise factors (seperti kelembaban, temperatur, dll) menunjukkan efek paling minimal terhadap nilai respon, menghasilkan proses yang lebih robust.

## Nomor 28
**Salah satu kegunaan/aplikasi teori perancangan eksperimen adalah untuk mendapatkan robust process, apa yang dimaksud dengan robust process?**

Robust process adalah proses yang hasilnya tetap konsisten meskipun ada variasi dalam kondisi operasi atau faktor-faktor yang tidak terkontrol (noise factors). Ciri-ciri robust process:

1. **Tidak sensitif terhadap noise factors** - Performa proses tetap stabil meskipun ada perubahan pada variabel yang tidak terkontrol seperti temperatur lingkungan, kelembaban, variasi bahan baku, dll.

2. **Memiliki variabilitas rendah** - Output proses memiliki varians yang kecil di sekitar nilai target.

3. **Dapat diprediksi** - Perilaku proses konsisten dan dapat diprediksi dari waktu ke waktu.

4. **Target performance tercapai** - Selain stabil, proses juga menghasilkan output yang sesuai dengan target kualitas yang diinginkan.

Dalam konteks perancangan eksperimen, robust process biasanya dicapai dengan metode Taguchi yang fokus pada meminimalkan sensitivitas terhadap noise factors sambil tetap mencapai target performa.

## Nomor 29
**Bagaimana kriteria sebuah eksperimen dapat dikatakan gagal/berhasil secara statistik atau praktis (statistical significance versus practical significance)? Apakah eksperimen pada skripsi tsb dapat dikatakan berhasil?**

### Statistical Significance vs Practical Significance

**Statistical Significance:**
- Menunjukkan bahwa hasil eksperimen cukup kuat secara statistik, biasanya ditunjukkan dengan p-value < α (level signifikansi)
- F-hitung > F-tabel dalam analisis ANOVA
- Menolak hipotesis nol (H0)
- Menunjukkan bahwa perubahan yang terdeteksi kemungkinan besar bukan karena kebetulan

**Practical Significance:**
- Menunjukkan bahwa hasil yang didapat memiliki nilai praktis
- Efek/perubahan yang ditemukan cukup besar untuk diimplementasikan
- Mempertimbangkan aspek ekonomi, teknis, dan operasional
- Memperhatikan ROI (Return on Investment) dari implementasi hasil

### Evaluasi Eksperimen pada Skripsi

Eksperimen pada skripsi tersebut dapat dikatakan berhasil, karena:

1. **Secara Statistik:**
   - Tiga faktor utama (Komposisi, Kompaksi, Cavity) terbukti signifikan dengan F-hitung > F-tabel
   - Hipotesis nol untuk ketiga faktor utama berhasil ditolak

2. **Secara Praktis:**
   - Kombinasi optimal yang ditemukan dapat diaplikasikan dalam produksi material penyerap suara
   - Peningkatan nilai decibel drop yang didapatkan cukup signifikan (>30%) dibanding baseline
   - Hasil dapat diimplementasikan untuk meningkatkan performa produk tanpa menambah biaya produksi secara signifikan

## Nomor 30
**Lampirkan kode python untuk men-generate tabel ANOVA, post-hoc test, dan kode untuk uji asumsi-asumsi anova (uji secara statistik). Tambahkan kode untuk men-generate grafik visual yang dibutuhkan (scatter plot kesamaan variansi, normalitas, independensi, dll).**

Kode Python untuk analisis ANOVA, post-hoc test dan uji asumsi telah dibuat dalam file `anova_analysis.py`. Berikut adalah penjelasan singkat mengenai bagian-bagian utama dari kode:

1. **Tabel ANOVA**
   - Menggunakan `statsmodels.formula.api.ols` dan `statsmodels.stats.anova.anova_lm`
   - Menguji signifikansi faktor utama dan interaksi

2. **Post-hoc Test**
   - Tukey HSD untuk perbandingan berpasangan
   - LSD (Least Significant Difference) untuk mengevaluasi perbedaan antar level faktor

3. **Uji Asumsi ANOVA**
   - Normalitas: Shapiro-Wilk test dan Q-Q plot
   - Homogenitas variansi: Levene's test dan scatter plot residual
   - Independensi: Durbin-Watson test dan plot residual vs urutan run

4. **Grafik Visual**
   - Main effects plot untuk setiap faktor
   - Interaction plots untuk interaksi antar faktor
   - Residual plots untuk evaluasi asumsi
   - Surface plots untuk visualisasi efek dua faktor terhadap respon

Silakan lihat file `anova_analysis.py` untuk implementasi lengkap dari kode analisis ini. 