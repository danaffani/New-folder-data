# Dokumentasi tabel_no3_tambahan.py

## Tujuan Program

Program `tabel_no3_tambahan.py` bertujuan untuk memproses dan mengorganisasi data eksperimen faktorial 3³ (tiga faktor dengan tiga level) dari file Excel sumber. Program ini mengolah dua jenis desain eksperimen:

1. **Randomized Block Design (RBD)**: Desain dengan blok-blok yang diatur berdasarkan replikasi
2. **Completely Randomized Design (CRD)**: Desain yang sepenuhnya teracak

Hasil dari program ini adalah file Excel yang berisi matriks desain yang telah dikode ulang dengan nilai numerik untuk memudahkan analisis statistik.

## Library yang Digunakan

Program ini menggunakan beberapa library Python:

- **pandas**: Untuk pengelolaan dan manipulasi data dalam format DataFrame serta mengekspor ke Excel
- **numpy**: Untuk operasi numerik dasar

## Struktur dan Fungsi Utama

Program terdiri dari satu fungsi utama `create_design_matrices()` yang melakukan beberapa tahapan pemrosesan:

### `create_design_matrices()`

Fungsi utama yang:

```python
def create_design_matrices():
    try:
        # Membaca data dari file Excel sumber
        excel_file = "input/tabel_koef_Serap_bunyi.xlsx"
        
        df_rbd = pd.read_excel(excel_file, sheet_name='RBD')
        df_crd = pd.read_excel(excel_file, sheet_name='CRD')
        
        # Mengkonversi nilai kategoris menjadi kode numerik
        for df in [df_rbd, df_crd]:
            df['A'] = df['Komposisi'].map({"50 : 50": -1, "70 : 30": 0, "90 : 10": 1})
            df['B'] = df['Kompaksi'].map({"3 : 4": -1, "4 : 4": 0, "5 : 4": 1})
            df['C'] = df['Cavity'].map({"15 mm": -1, "20 mm": 0, "25 mm": 1})
        
        # Menyimpan hasil ke Excel
        with pd.ExcelWriter('output/nomor3_tambahan.xlsx') as writer:
            # Sheet untuk RBD dan CRD
            df_rbd.to_excel(writer, sheet_name='RBD', index=False)
            df_crd.to_excel(writer, sheet_name='CRD', index=False)
            
            # Membuat ringkasan data respon untuk nilai rata-rata
            df_rata = df_rbd[df_rbd['Frekuensi'] == "Rata-rata"].copy()
            
            # Menghitung mean respon untuk setiap kombinasi faktor
            response_data = pd.DataFrame({
                'A': [-1, -1, -1, -1, 1, 1, 1, 1],
                'B': [-1, -1, 1, 1, -1, -1, 1, 1],
                'C': [-1, 1, -1, 1, -1, 1, -1, 1],
                'Response_Mean': [df_rata[
                    (df_rata['A'] == a) & 
                    (df_rata['B'] == b) & 
                    (df_rata['C'] == c)
                ]['Rata-rata Koefisien'].mean() 
                for a, b, c in zip(
                    [-1, -1, -1, -1, 1, 1, 1, 1],
                    [-1, -1, 1, 1, -1, -1, 1, 1],
                    [-1, 1, -1, 1, -1, 1, -1, 1]
                )]
            })
            response_data.to_excel(writer, sheet_name='Response_Data', index=False)
        
        print("Design matrices telah dibuat dan disimpan di output/nomor3_tambahan.xlsx")
        print(f"Jumlah baris RBD: {len(df_rbd)} (dengan semua frekuensi dan rata-rata)")
        print(f"Jumlah baris CRD: {len(df_crd)} (dengan semua frekuensi dan rata-rata)")
        print("Data diambil dari file tabel_koef_Serap_bunyi.xlsx")
        
    except Exception as e:
        print(f"Terjadi error: {str(e)}")
```

Proses kerjanya:

1. Membaca data dari file Excel sumber (`tabel_koef_Serap_bunyi.xlsx`)
2. Mengkonversi nilai kategoris untuk faktor-faktor (Komposisi, Kompaksi, Cavity) menjadi kode numerik:
   - Komposisi: "50 : 50" → -1, "70 : 30" → 0, "90 : 10" → 1
   - Kompaksi: "3 : 4" → -1, "4 : 4" → 0, "5 : 4" → 1
   - Cavity: "15 mm" → -1, "20 mm" → 0, "25 mm" → 1
3. Menyimpan hasil konversi ke file Excel baru (`output/nomor3_tambahan.xlsx`)
4. Membuat tabel tambahan yang berisi ringkasan nilai respon rata-rata untuk setiap kombinasi faktor

## Desain Eksperimen yang Diproses

Program memproses desain faktorial 3³, yang berarti eksperimen melibatkan 3 faktor (Komposisi, Kompaksi, Cavity) dengan masing-masing 3 level. Total kombinasi perlakuan adalah 3³ = 27 kombinasi.

### Replikasi dan Blocks

Dalam data eksperimental yang diproses:

1. **Replikasi**: Eksperimen dilakukan dengan 3 replikasi (pengulangan), yang direpresentasikan oleh kolom `Blocks` dengan nilai 1, 2, dan 3.
   - Setiap replikasi mencakup semua 27 kombinasi perlakuan
   - Replikasi digunakan untuk meningkatkan keakuratan estimasi dan mengurangi variabilitas

2. **Total pengamatan**: 27 kombinasi perlakuan × 3 replikasi = 81 pengamatan per frekuensi

Koding numerik yang digunakan:

| Faktor       | Level Kategoris | Kode Numerik |
|--------------|-----------------|--------------|
| Komposisi    | 50 : 50         | -1           |
|              | 70 : 30         | 0            |
|              | 90 : 10         | 1            |
| Kompaksi     | 3 : 4           | -1           |
|              | 4 : 4           | 0            |
|              | 5 : 4           | 1            |
| Cavity       | 15 mm           | -1           |
|              | 20 mm           | 0            |
|              | 25 mm           | 1            |

### Randomized Block Design (RBD) vs Completely Randomized Design (CRD)

Program ini memproses data yang sudah diatur dalam dua jenis desain eksperimen:

1. **RBD (Randomized Block Design)**: 
   - Eksperimen dibagi menjadi 3 blok (replikasi) dengan nilai Blocks = 1, 2, dan 3
   - Setiap blok berisi semua 27 kombinasi perlakuan
   - Pengacakan dilakukan dalam setiap blok
   - Memungkinkan untuk mengendalikan pengaruh blok (misalnya waktu eksperimen atau batch spesimen) yang dapat mempengaruhi hasil

2. **CRD (Completely Randomized Design)**:
   - Semua pengamatan berada dalam satu kelompok (Blocks = 1)
   - Pengacakan dilakukan untuk seluruh eksperimen tanpa mempertimbangkan blok
   - Lebih sederhana, tetapi tidak memperhitungkan pengaruh blok

Perbedaan utama adalah dalam RBD, perbedaan antar blok (misalnya variasi antar hari pengujian atau batch material) dapat diidentifikasi dan dihilangkan dari error, sehingga meningkatkan sensitivitas analisis.

## Output Program

Program menghasilkan file Excel 'output/nomor3_tambahan.xlsx' dengan tiga sheet:

1. **RBD**: Matriks desain untuk Randomized Block Design dengan pengkodean numerik
2. **CRD**: Matriks desain untuk Completely Randomized Design dengan pengkodean numerik
3. **Response_Data**: Ringkasan nilai respon rata-rata untuk kombinasi faktor tertentu

Contoh struktur data RBD/CRD yang dihasilkan:

| Komposisi | Kompaksi | Cavity | Blocks | A  | B  | C  | Rata-rata Koefisien | ... |
|-----------|----------|--------|--------|----|----|----|---------------------|-----|
| 50 : 50   | 3 : 4    | 15 mm  | 1      | -1 | -1 | -1 | 0.21                | ... |
| 50 : 50   | 3 : 4    | 20 mm  | 1      | -1 | -1 | 0  | 0.23                | ... |
| 50 : 50   | 3 : 4    | 25 mm  | 1      | -1 | -1 | 1  | 0.25                | ... |
| ...       | ...      | ...    | ...    | ...| ...| ...| ...                 | ... |
| 50 : 50   | 3 : 4    | 15 mm  | 2      | -1 | -1 | -1 | 0.20                | ... |
| ...       | ...      | ...    | ...    | ...| ...| ...| ...                 | ... |
| 50 : 50   | 3 : 4    | 15 mm  | 3      | -1 | -1 | -1 | 0.22                | ... |
| ...       | ...      | ...    | ...    | ...| ...| ...| ...                 | ... |

## Penggunaan Lebih Lanjut

Data yang diproses oleh program ini dapat digunakan untuk:

1. Analisis ANOVA faktorial untuk menentukan pengaruh faktor dan interaksinya
2. Pembuatan plot interaksi antar faktor
3. Analisis respons permukaan untuk optimasi
4. Pemodelan statistik lanjutan

Program ini menjadi proses perantara yang penting dalam rantai analisis, karena mengkonversi data kategoris menjadi kode numerik yang lebih mudah diproses oleh algoritma statistik. 