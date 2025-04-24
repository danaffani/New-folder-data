import pandas as pd
import numpy as np

def create_design_matrices():
    try:
        excel_file = "input/tabel_koef_Serap_bunyi.xlsx"
        
        df_rbd = pd.read_excel(excel_file, sheet_name='RBD')
        df_crd = pd.read_excel(excel_file, sheet_name='CRD')
        
        for df in [df_rbd, df_crd]:
            df['A'] = df['Komposisi'].map({"50 : 50": -1, "70 : 30": 0, "90 : 10": 1})
            df['B'] = df['Kompaksi'].map({"3 : 4": -1, "4 : 4": 0, "5 : 4": 1})
            df['C'] = df['Cavity'].map({"15 mm": -1, "20 mm": 0, "25 mm": 1})
        
        with pd.ExcelWriter('output/nomor3_tambahan.xlsx') as writer:
            df_rbd.to_excel(writer, sheet_name='RBD', index=False)
            df_crd.to_excel(writer, sheet_name='CRD', index=False)
            
            df_rata = df_rbd[df_rbd['Frekuensi'] == "Rata-rata"].copy()
            
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

if __name__ == "__main__":
    create_design_matrices() 