import pandas as pd
import numpy as np

def create_design_matrices():
    # Baca file Excel
    try:
        excel_file = "input/semua_tabel.xlsx"
        data = pd.read_excel(excel_file, sheet_name=None)
        
        def calculate_response(A, B, C):
            # Perhitungan nilai respon berdasarkan kombinasi faktor
            # Rumus ini bisa disesuaikan dengan kebutuhan
            base_value = 50  # Nilai dasar
            
            # Efek utama
            effect_A = 5 * A  # Pengaruh faktor A
            effect_B = 3 * B  # Pengaruh faktor B
            effect_C = 4 * C  # Pengaruh faktor C
            
            # Interaksi dua faktor
            interaction_AB = 2 * A * B
            interaction_AC = 1.5 * A * C
            interaction_BC = 1 * B * C
            
            # Interaksi tiga faktor
            interaction_ABC = 0.5 * A * B * C
            
            # Random noise untuk variasi
            noise = np.random.normal(0, 1)
            
            # Total nilai respon
            response = base_value + effect_A + effect_B + effect_C + \
                      interaction_AB + interaction_AC + interaction_BC + \
                      interaction_ABC + noise
            
            return round(response, 2)

        # Fungsi untuk membuat design matrix
        def create_matrix(design_type):
            n_treatments = 8
            n_replications = 3
            total_runs = n_treatments * n_replications
            
            # Buat dataframe dasar
            df = pd.DataFrame({
                'StdOrder': range(1, total_runs + 1),
                'Blocks': [i for i in range(1, n_replications + 1) for _ in range(n_treatments)] if design_type == 'RBD' else [1] * total_runs,
                'A': [-1, -1, -1, -1, 1, 1, 1, 1] * n_replications,
                'B': [-1, -1, 1, 1, -1, -1, 1, 1] * n_replications,
                'C': [-1, 1, -1, 1, -1, 1, -1, 1] * n_replications
            })
            
            # Generate RunOrder
            if design_type == 'RBD':
                # For RBD, randomize within each block
                run_order = []
                for block in range(1, n_replications + 1):
                    block_indices = list(range((block-1)*n_treatments + 1, block*n_treatments + 1))
                    np.random.shuffle(block_indices)
                    run_order.extend(block_indices)
            else:
                # For CRD, completely randomize all runs
                run_order = list(range(1, total_runs + 1))
                np.random.shuffle(run_order)
            
            df['RunOrder'] = run_order
            
            # Sort by RunOrder untuk mendapatkan urutan eksekusi
            df = df.sort_values('RunOrder').reset_index(drop=True)
            
            # Generate jadwal running (contoh: setiap 30 menit)
            start_time = pd.Timestamp('2024-01-01 08:00:00')
            df['Jadwal'] = [start_time + pd.Timedelta(minutes=30*i) for i in range(len(df))]
            df['Jadwal'] = df['Jadwal'].dt.strftime('%H:%M')
            
            # Hitung nilai respon untuk setiap kombinasi faktor
            df['Nilai Respon'] = df.apply(lambda row: calculate_response(row['A'], row['B'], row['C']), axis=1)
            
            return df
        
        # Set random seed untuk reproducibility
        np.random.seed(42)
        
        # Buat kedua jenis design matrix
        rbd_matrix = create_matrix('RBD')
        crd_matrix = create_matrix('CRD')
        
        # Simpan ke Excel
        with pd.ExcelWriter('output/nomor3_tambahan.xlsx') as writer:
            rbd_matrix.to_excel(writer, sheet_name='RBD', index=False)
            crd_matrix.to_excel(writer, sheet_name='CRD', index=False)
            
            # Simpan juga data respon untuk digunakan di analisis ANOVA
            response_data = pd.DataFrame({
                'A': [-1, -1, -1, -1, 1, 1, 1, 1],
                'B': [-1, -1, 1, 1, -1, -1, 1, 1],
                'C': [-1, 1, -1, 1, -1, 1, -1, 1],
                'Response_Mean': [rbd_matrix[
                    (rbd_matrix['A'] == a) & 
                    (rbd_matrix['B'] == b) & 
                    (rbd_matrix['C'] == c)
                ]['Nilai Respon'].mean() 
                for a, b, c in zip(
                    [-1, -1, -1, -1, 1, 1, 1, 1],
                    [-1, -1, 1, 1, -1, -1, 1, 1],
                    [-1, 1, -1, 1, -1, 1, -1, 1]
                )]
            })
            response_data.to_excel(writer, sheet_name='Response_Data', index=False)
        
        print("Design matrices telah dibuat dan disimpan di output/nomor3_tambahan.xlsx")
        
    except Exception as e:
        print(f"Terjadi error: {str(e)}")

if __name__ == "__main__":
    create_design_matrices() 