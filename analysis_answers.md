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
**Kode Python untuk Analisis**

Implementasi lengkap dapat dilihat di file:
1. `tabel_no3_tambahan.py` - Design matrix dan perhitungan nilai respon
2. `tabel_no8_tambahan.py` - Analisis ANOVA dan perbandingan hasil

Detail implementasi dan penjelasan kode dapat dilihat di bagian [Nomor 3 Tambahan](#nomor-3-tambahan) dan [Nomor 8 Tambahan](#nomor-8-tambahan) di atas. 