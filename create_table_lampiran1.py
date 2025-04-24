import pandas as pd
import numpy as np
import os

def buat_tabel_lampiran1():
    """
    Fungsi utama untuk membuat tabel lampiran 1 dari data koefisien serap bunyi.
    """
    # Membuat direktori output jika belum ada
    if not os.path.exists('output'):
        os.makedirs('output')
    
    # Membaca data dari file Excel
    file_path = 'input/tabel_koef_Serap_bunyi_revised.xlsx'
    try:
        df_rbd = pd.read_excel(file_path, sheet_name='RBD')
        print(f"Data berhasil dibaca dengan shape: {df_rbd.shape}")
    except Exception as e:
        print(f"Error membaca file Excel: {e}")
        return
    
    # Membuat tabel koefisien serap bunyi per frekuensi
    buat_tabel_koefisien_serap_bunyi_per_frekuensi(df_rbd)
    
    # Membuat tabel RBD dan CRD
    buat_tabel_rbd_crd(df_rbd)
    
    print("Pembuatan tabel selesai!")

def buat_tabel_koefisien_serap_bunyi_per_frekuensi(df_rbd):
    """
    Membuat tabel koefisien serap bunyi per frekuensi berdasarkan data RBD.
    
    Parameters:
    df_rbd (DataFrame): Data dari sheet RBD
    """
    # Menyaring data untuk masing-masing frekuensi
    frekuensi_list = df_rbd['Frekuensi'].unique()
    
    # Menyiapkan DataFrame untuk menyimpan hasil
    results = []
    
    for frekuensi in frekuensi_list:
        df_freq = df_rbd[df_rbd['Frekuensi'] == frekuensi]
        
        # Mengelompokkan berdasarkan kombinasi faktor
        grouped = df_freq.groupby(['Komposisi', 'Kompaksi', 'Cavity'])
        
        for name, group in grouped:
            komposisi, kompaksi, cavity = name
            
            # Dapatkan nilai spesimen untuk setiap blocks (1-5)
            specimens = []
            for block in range(1, 6):
                nilai = group[group['Blocks'] == block]['Nilai Spesimen'].values
                if len(nilai) > 0:
                    specimens.append(nilai[0])
                else:
                    specimens.append(np.nan)
            
            # Hitung rata-rata dan standard deviasi
            mean = np.mean(specimens)
            std_dev = np.std(specimens)
            
            # Tambahkan ke hasil
            results.append({
                'Frekuensi': frekuensi,
                'Komposisi': komposisi,
                'Kompaksi': kompaksi, 
                'Cavity': cavity,
                'Spesimen 1': specimens[0],
                'Spesimen 2': specimens[1],
                'Spesimen 3': specimens[2],
                'Spesimen 4': specimens[3],
                'Spesimen 5': specimens[4],
                'Rata-rata': mean,
                'Standar Deviasi': std_dev
            })
    
    # Buat DataFrame dari hasil
    result_df = pd.DataFrame(results)
    
    # Urutkan hasil
    result_df = result_df.sort_values(by=['Frekuensi', 'Komposisi', 'Kompaksi', 'Cavity'])
    
    # Simpan hasil ke Excel
    output_path = 'output/tabel_koefisien_serap_bunyi_per_frekuensi.xlsx'
    result_df.to_excel(output_path, index=False)
    print(f"Tabel koefisien serap bunyi per frekuensi telah disimpan di: {output_path}")

def buat_tabel_rbd_crd(df_rbd):
    """
    Membuat tabel Randomized Block Design (RBD) dan Completely Randomized Design (CRD).
    
    Parameters:
    df_rbd (DataFrame): Data dari sheet RBD
    """
    # Menyalin data RBD
    rbd_df = df_rbd.copy()
    
    # Membuat data CRD (dengan blocks dirandom)
    crd_df = df_rbd.copy()
    
    # Mengacak run order untuk CRD
    crd_df['RunOrder'] = np.random.permutation(crd_df['RunOrder'].values)
    crd_df = crd_df.sort_values('RunOrder')
    
    # Menyimpan kedua tabel ke Excel
    with pd.ExcelWriter('output/tabel_rbd_crd.xlsx') as writer:
        rbd_df.to_excel(writer, sheet_name='RBD', index=False)
        crd_df.to_excel(writer, sheet_name='CRD', index=False)
    
    print("Tabel RBD dan CRD telah disimpan di: output/tabel_rbd_crd.xlsx")

if __name__ == "__main__":
    buat_tabel_lampiran1() 