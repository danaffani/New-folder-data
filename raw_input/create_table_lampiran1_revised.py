import pandas as pd
import numpy as np
import os

def buat_tabel_koefisien_serap_bunyi_per_frekuensi():
    data = [
        ["50 : 50", "3 : 4", "15 mm", 0.036, 0.144, 0.376, 0.893, 0.037, 0.133, 0.388, 0.902, 0.043, 0.131, 0.391, 0.884, 0.039, 0.144, 0.377, 0.897, 0.039, 0.151, 0.375, 0.881],
        ["50 : 50", "3 : 4", "20 mm", 0.028, 0.097, 0.218, 0.799, 0.029, 0.096, 0.221, 0.815, 0.035, 0.103, 0.236, 0.805, 0.031, 0.100, 0.220, 0.799, 0.035, 0.105, 0.223, 0.807],
        ["50 : 50", "3 : 4", "25 mm", 0.038, 0.106, 0.243, 0.835, 0.041, 0.104, 0.241, 0.821, 0.041, 0.105, 0.247, 0.841, 0.038, 0.116, 0.255, 0.822, 0.038, 0.118, 0.255, 0.829],
        ["50 : 50", "4 : 4", "15 mm", 0.036, 0.196, 0.355, 0.744, 0.036, 0.202, 0.346, 0.741, 0.038, 0.201, 0.370, 0.754, 0.036, 0.199, 0.366, 0.758, 0.036, 0.203, 0.366, 0.753],
        ["50 : 50", "4 : 4", "20 mm", 0.038, 0.166, 0.254, 0.646, 0.036, 0.150, 0.269, 0.654, 0.038, 0.166, 0.264, 0.667, 0.037, 0.169, 0.262, 0.642, 0.036, 0.151, 0.246, 0.662],
        ["50 : 50", "4 : 4", "25 mm", 0.034, 0.161, 0.250, 0.660, 0.035, 0.162, 0.266, 0.666, 0.036, 0.168, 0.254, 0.668, 0.035, 0.156, 0.263, 0.661, 0.034, 0.156, 0.267, 0.665],
        ["50 : 50", "5 : 4", "15 mm", 0.035, 0.248, 0.334, 0.595, 0.035, 0.258, 0.337, 0.600, 0.038, 0.254, 0.333, 0.594, 0.033, 0.255, 0.333, 0.605, 0.030, 0.258, 0.346, 0.597],
        ["50 : 50", "5 : 4", "20 mm", 0.047, 0.234, 0.289, 0.492, 0.056, 0.245, 0.285, 0.505, 0.049, 0.227, 0.295, 0.503, 0.056, 0.242, 0.287, 0.502, 0.047, 0.245, 0.292, 0.496],
        ["50 : 50", "5 : 4", "25 mm", 0.029, 0.216, 0.256, 0.485, 0.033, 0.206, 0.256, 0.485, 0.031, 0.219, 0.253, 0.485, 0.028, 0.215, 0.266, 0.496, 0.028, 0.211, 0.253, 0.490],
        ["70 : 30", "3 : 4", "15 mm", 0.008, 0.203, 0.516, 0.831, 0.010, 0.200, 0.518, 0.839, 0.005, 0.211, 0.509, 0.834, 0.019, 0.204, 0.515, 0.829, 0.009, 0.210, 0.509, 0.839],
        ["70 : 30", "3 : 4", "20 mm", 0.023, 0.104, 0.237, 0.916, 0.024, 0.107, 0.244, 0.876, 0.026, 0.102, 0.232, 0.882, 0.026, 0.104, 0.247, 0.878, 0.026, 0.113, 0.231, 0.874],
        ["70 : 30", "3 : 4", "25 mm", 0.021, 0.106, 0.215, 0.779, 0.040, 0.108, 0.228, 0.800, 0.032, 0.112, 0.213, 0.776, 0.034, 0.112, 0.217, 0.773, 0.033, 0.107, 0.225, 0.783],
        ["70 : 30", "4 : 4", "15 mm", 0.025, 0.258, 0.462, 0.775, 0.027, 0.213, 0.507, 0.774, 0.031, 0.234, 0.502, 0.771, 0.022, 0.254, 0.500, 0.785, 0.030, 0.261, 0.458, 0.788],
        ["70 : 30", "4 : 4", "20 mm", 0.030, 0.190, 0.299, 0.815, 0.025, 0.203, 0.283, 0.819, 0.049, 0.207, 0.301, 0.819, 0.031, 0.203, 0.310, 0.813, 0.049, 0.201, 0.280, 0.810],
        ["70 : 30", "4 : 4", "25 mm", 0.017, 0.177, 0.295, 0.784, 0.014, 0.164, 0.282, 0.776, 0.022, 0.167, 0.304, 0.786, 0.013, 0.177, 0.288, 0.786, 0.029, 0.180, 0.284, 0.799],
        ["70 : 30", "5 : 4", "15 mm", 0.041, 0.313, 0.471, 0.719, 0.050, 0.329, 0.475, 0.725, 0.054, 0.325, 0.463, 0.716, 0.048, 0.324, 0.472, 0.717, 0.043, 0.313, 0.471, 0.723],
        ["70 : 30", "5 : 4", "20 mm", 0.036, 0.276, 0.361, 0.713, 0.050, 0.284, 0.372, 0.715, 0.033, 0.283, 0.370, 0.719, 0.034, 0.281, 0.366, 0.727, 0.044, 0.266, 0.371, 0.728],
        ["70 : 30", "5 : 4", "25 mm", 0.013, 0.248, 0.374, 0.789, 0.026, 0.244, 0.384, 0.780, 0.024, 0.259, 0.384, 0.793, 0.024, 0.244, 0.377, 0.792, 0.012, 0.266, 0.373, 0.791],
        ["90 : 10", "3 : 4", "15 mm", 0.021, 0.301, 0.563, 0.647, 0.023, 0.319, 0.571, 0.638, 0.036, 0.301, 0.564, 0.636, 0.035, 0.318, 0.568, 0.631, 0.035, 0.318, 0.567, 0.645],
        ["90 : 10", "3 : 4", "20 mm", 0.019, 0.244, 0.473, 0.796, 0.010, 0.239, 0.478, 0.789, 0.027, 0.238, 0.470, 0.799, 0.030, 0.255, 0.472, 0.781, 0.014, 0.248, 0.476, 0.789],
        ["90 : 10", "3 : 4", "25 mm", 0.016, 0.196, 0.409, 0.799, 0.028, 0.196, 0.401, 0.804, 0.016, 0.188, 0.412, 0.801, 0.033, 0.185, 0.409, 0.817, 0.035, 0.207, 0.411, 0.807],
        ["90 : 10", "4 : 4", "15 mm", 0.023, 0.249, 0.420, 0.576, 0.037, 0.246, 0.414, 0.578, 0.014, 0.257, 0.436, 0.583, 0.023, 0.251, 0.429, 0.574, 0.026, 0.241, 0.434, 0.577],
        ["90 : 10", "4 : 4", "20 mm", 0.033, 0.215, 0.343, 0.573, 0.033, 0.206, 0.352, 0.583, 0.037, 0.217, 0.344, 0.587, 0.035, 0.228, 0.352, 0.578, 0.021, 0.208, 0.340, 0.573],
        ["90 : 10", "4 : 4", "25 mm", 0.022, 0.195, 0.310, 0.603, 0.034, 0.188, 0.310, 0.611, 0.036, 0.181, 0.318, 0.606, 0.033, 0.204, 0.311, 0.616, 0.015, 0.192, 0.314, 0.616],
        ["90 : 10", "5 : 4", "15 mm", 0.024, 0.196, 0.277, 0.504, 0.019, 0.189, 0.287, 0.515, 0.018, 0.199, 0.274, 0.502, 0.020, 0.189, 0.274, 0.498, 0.034, 0.200, 0.287, 0.492],
        ["90 : 10", "5 : 4", "20 mm", 0.047, 0.185, 0.213, 0.349, 0.047, 0.198, 0.215, 0.356, 0.048, 0.207, 0.211, 0.358, 0.034, 0.194, 0.201, 0.359, 0.040, 0.181, 0.215, 0.340],
        ["90 : 10", "5 : 4", "25 mm", 0.027, 0.193, 0.211, 0.407, 0.036, 0.199, 0.220, 0.416, 0.030, 0.209, 0.216, 0.408, 0.036, 0.195, 0.203, 0.412, 0.032, 0.184, 0.224, 0.419]
    ]
    
    processed_data = []
    frekuensi_names = ["150 Hz", "300 Hz", "600 Hz", "1200 Hz"]
    
    for row in data:
        komposisi, kompaksi, cavity = row[0], row[1], row[2]
        values = row[3:]
        
        for freq_idx, freq_name in enumerate(frekuensi_names):
            specimen_values = []
            
            for i in range(5):
                start_idx = i * 4 + freq_idx
                specimen_values.append(values[start_idx])
            
            avg = round(sum(specimen_values) / len(specimen_values), 3)
            
            processed_data.append([komposisi, kompaksi, cavity, freq_name] + specimen_values + [avg])
    
    return processed_data

def tambahkan_rata_rata_frekuensi(data):
    factor_data = {}
    
    for row in data:
        komposisi, kompaksi, cavity = row[0], row[1], row[2]
        key = (komposisi, kompaksi, cavity)
        
        if key not in factor_data:
            factor_data[key] = []
        
        factor_data[key].append(row)
    
    avg_data = []
    
    for key, rows in factor_data.items():
        komposisi, kompaksi, cavity = key
        
        spec_avgs = [0.0, 0.0, 0.0, 0.0, 0.0]
        
        for row in rows:
            for i in range(5):
                spec_avgs[i] += row[4+i] / 4
        
        spec_avgs = [round(val, 3) for val in spec_avgs]
        
        overall_avg = round(sum(spec_avgs) / len(spec_avgs), 3)
        
        avg_data.append([komposisi, kompaksi, cavity, "Rata-rata"] + spec_avgs + [overall_avg])
    
    return data + avg_data

def buat_tabel_rbd_crd(data):
    """
    Fungsi yang direvisi untuk membuat tabel RBD dan CRD.
    - RBD: Menggunakan blocks 1-5 sesuai dengan spesimen
    - CRD: Struktur yang sama dengan blocks = spesimen
    """
    rbd_data = []
    crd_data = []
    
    # Membuat daftar kombinasi faktor
    combinations = []
    for item in data:
        komposisi, kompaksi, cavity, frekuensi = item[0], item[1], item[2], item[3]
        combinations.append((komposisi, kompaksi, cavity, frekuensi))
    
    np.random.seed(42)
    std_order = 1
    
    # Membuat data RBD dengan block 1-5 (representasi spesimen)
    for combination_idx, combination in enumerate(combinations):
        komposisi, kompaksi, cavity, frekuensi = combination
        specimen_values = data[combination_idx][4:9]  # Nilai untuk 5 spesimen
        overall_avg = data[combination_idx][9]
        
        # Untuk setiap spesimen, kita membuat baris dalam RBD
        for specimen in range(5):
            block = specimen + 1  # Block 1-5
            specimen_value = specimen_values[specimen]
            run_order = std_order
            
            rbd_data.append([
                std_order, 
                run_order, 
                block,  # Block 1-5
                komposisi, 
                kompaksi, 
                cavity, 
                frekuensi, 
                specimen_value,  # Nilai untuk spesimen ini
                overall_avg
            ])
            
            # Juga tambahkan ke CRD
            crd_data.append([
                std_order, 
                run_order, 
                block,  # Block = spesimen
                komposisi, 
                kompaksi, 
                cavity, 
                frekuensi, 
                specimen_value,  # Nilai untuk spesimen ini  
                overall_avg
            ])
            
            std_order += 1
    
    # Membuat DataFrame
    rbd_df = pd.DataFrame(rbd_data, columns=[
        'StdOrder', 'RunOrder', 'Blocks', 'Komposisi', 'Kompaksi', 'Cavity', 'Frekuensi',
        'Nilai Spesimen', 'Rata-rata Koefisien'
    ])
    
    crd_df = pd.DataFrame(crd_data, columns=[
        'StdOrder', 'RunOrder', 'Blocks', 'Komposisi', 'Kompaksi', 'Cavity', 'Frekuensi',
        'Nilai Spesimen', 'Rata-rata Koefisien'
    ])
    
    return rbd_df, crd_df

def main():
    if not os.path.exists('input'):
        os.makedirs('input')
    
    data = buat_tabel_koefisien_serap_bunyi_per_frekuensi()
    
    data_with_avg = tambahkan_rata_rata_frekuensi(data)
    
    rbd_df, crd_df = buat_tabel_rbd_crd(data_with_avg)
    
    with pd.ExcelWriter('input/tabel_koef_Serap_bunyi.xlsx') as writer:
        rbd_df.to_excel(writer, sheet_name='RBD', index=False)
        crd_df.to_excel(writer, sheet_name='CRD', index=False)
    
    print("File Excel tabel_koef_Serap_bunyi.xlsx berhasil dibuat di folder input/")
    print(f"File berisi sheet RBD dengan {len(rbd_df)} baris dan CRD dengan {len(crd_df)} baris")
    print("Data sudah dipisahkan per frekuensi dan ditambahkan rata-rata frekuensi")

if __name__ == "__main__":
    main() 