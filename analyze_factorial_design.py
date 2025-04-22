import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import statsmodels.api as sm
from statsmodels.formula.api import ols
from statsmodels.stats.anova import anova_lm
import seaborn as sns
from itertools import combinations
import glob
import os
from scipy import stats

# Fungsi untuk membaca data dari file txt
def read_data_from_txt():
    # Membaca data dari tabel_4.1.txt (NRC) dan tabel_4.2.txt (decibel drop)
    nrc_file = "tabel_4.1.txt"
    decibel_file = "tabel_4.2.txt"
    
    # Membaca data NRC
    with open(nrc_file, 'r') as file:
        nrc_lines = file.readlines()
    
    # Membaca data decibel drop
    with open(decibel_file, 'r') as file:
        decibel_lines = file.readlines()
    
    # Parsing data ke dalam format yang dapat dianalisis
    compositions = ["50:50", "70:30", "90:10"]
    compactions = ["3:4", "4:4", "5:4"]
    cavities = ["15mm", "20mm", "25mm"]
    
    # Struktur data untuk menyimpan semua nilai
    nrc_data = []
    decibel_data = []
    
    for comp_idx, comp in enumerate(compositions):
        for compac_idx, compac in enumerate(compactions):
            for cav_idx, cav in enumerate(cavities):
                # Menghitung indeks baris untuk data ini dalam file teks
                row_offset = comp_idx * 9 + compac_idx * 3 + cav_idx + 7  # Sesuaikan dengan format file
                
                # Ekstrak nilai NRC (5 replikasi)
                nrc_values = []
                for i in range(5):
                    try:
                        value = float(nrc_lines[row_offset].split('|')[i+2].strip())
                        nrc_values.append(value)
                    except:
                        nrc_values.append(np.nan)
                
                # Ekstrak nilai decibel drop (5 replikasi)
                decibel_values = []
                for i in range(5):
                    try:
                        value = float(decibel_lines[row_offset].split('|')[i+2].strip())
                        decibel_values.append(value)
                    except:
                        decibel_values.append(np.nan)
                
                # Menambahkan data ke daftar utama
                for rep in range(5):
                    nrc_data.append({
                        'Komposisi': comp,
                        'Kompaksi': compac,
                        'Cavity': cav,
                        'Replikasi': rep + 1,
                        'NRC': nrc_values[rep]
                    })
                    
                    decibel_data.append({
                        'Komposisi': comp,
                        'Kompaksi': compac,
                        'Cavity': cav,
                        'Replikasi': rep + 1,
                        'DecibelDrop': decibel_values[rep]
                    })
    
    # Membuat DataFrame
    nrc_df = pd.DataFrame(nrc_data)
    decibel_df = pd.DataFrame(decibel_data)
    
    return nrc_df, decibel_df

# Fungsi untuk melakukan analisis ANOVA
def analyze_anova(data, response_var):
    # Menentukan formula untuk ANOVA
    formula = f"{response_var} ~ C(Komposisi) + C(Kompaksi) + C(Cavity) + C(Komposisi):C(Kompaksi) + C(Komposisi):C(Cavity) + C(Kompaksi):C(Cavity) + C(Komposisi):C(Kompaksi):C(Cavity)"
    
    # Melakukan ANOVA
    model = ols(formula, data=data).fit()
    anova_result = anova_lm(model)
    
    return anova_result

# Fungsi untuk membuat plot interaksi
def create_interaction_plots(data, response_var, output_dir="plots"):
    # Memastikan direktori output ada
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Membuat DataFrame agregat untuk plotting
    agg_data = data.groupby(['Komposisi', 'Kompaksi', 'Cavity'])[response_var].mean().reset_index()
    
    # Plot interaksi Komposisi * Kompaksi
    plt.figure(figsize=(10, 6))
    for komp in agg_data['Komposisi'].unique():
        subset = agg_data[agg_data['Komposisi'] == komp]
        plt.plot(subset['Kompaksi'], subset[response_var], marker='o', label=komp)
    
    plt.xlabel('Kompaksi')
    plt.ylabel(response_var)
    plt.title(f'Interaksi Komposisi * Kompaksi untuk {response_var}')
    plt.legend()
    plt.grid(True)
    plt.savefig(f"{output_dir}/interaksi_komposisi_kompaksi_{response_var}.png")
    
    # Plot interaksi Komposisi * Cavity
    plt.figure(figsize=(10, 6))
    for komp in agg_data['Komposisi'].unique():
        subset = agg_data[agg_data['Komposisi'] == komp]
        plt.plot(subset['Cavity'], subset[response_var], marker='o', label=komp)
    
    plt.xlabel('Cavity')
    plt.ylabel(response_var)
    plt.title(f'Interaksi Komposisi * Cavity untuk {response_var}')
    plt.legend()
    plt.grid(True)
    plt.savefig(f"{output_dir}/interaksi_komposisi_cavity_{response_var}.png")
    
    # Plot interaksi Kompaksi * Cavity
    plt.figure(figsize=(10, 6))
    for komp in agg_data['Kompaksi'].unique():
        subset = agg_data[agg_data['Kompaksi'] == komp]
        plt.plot(subset['Cavity'], subset[response_var], marker='o', label=komp)
    
    plt.xlabel('Cavity')
    plt.ylabel(response_var)
    plt.title(f'Interaksi Kompaksi * Cavity untuk {response_var}')
    plt.legend()
    plt.grid(True)
    plt.savefig(f"{output_dir}/interaksi_kompaksi_cavity_{response_var}.png")

# Fungsi untuk melakukan analisis LSD (Least Significant Difference)
def perform_lsd_analysis(data, response_var, alpha=0.05):
    # Menghitung LSD untuk setiap faktor
    results = {}
    
    # Menghitung mean untuk setiap level faktor
    factor_means = {}
    for factor in ['Komposisi', 'Kompaksi', 'Cavity']:
        factor_means[factor] = data.groupby([factor])[response_var].mean()
    
    # Menghitung MSE dan df dari ANOVA
    anova_result = analyze_anova(data, response_var)
    mse = anova_result.loc['Residual', 'mean_sq']
    df_error = anova_result.loc['Residual', 'df']
    
    # Menghitung nilai t-kritik
    t_critical = stats.t.ppf(1 - alpha/2, df_error)
    
    # Menghitung LSD untuk setiap faktor
    for factor, means in factor_means.items():
        n = data.groupby([factor]).size().iloc[0]  # Jumlah observasi per level
        lsd = t_critical * np.sqrt(2 * mse / n)
        
        # Mengurutkan means
        sorted_means = means.sort_values()
        
        # Membuat pasangan perbandingan
        comparisons = []
        for i, j in combinations(sorted_means.index, 2):
            diff = abs(sorted_means[j] - sorted_means[i])
            significant = 'Ya' if diff > lsd else 'Tidak'
            comparisons.append({
                'Factor': factor,
                'Level 1': i,
                'Level 2': j,
                'Mean 1': sorted_means[i],
                'Mean 2': sorted_means[j],
                'Difference': diff,
                'LSD': lsd,
                'Significant': significant
            })
        
        results[factor] = pd.DataFrame(comparisons)
    
    return results

# Menjalankan semua analisis
if __name__ == "__main__":
    # Membaca data
    nrc_df, decibel_df = read_data_from_txt()
    
    # Menyimpan data ke Excel untuk verifikasi
    with pd.ExcelWriter("data_for_analysis.xlsx") as writer:
        nrc_df.to_excel(writer, sheet_name="NRC Data", index=False)
        decibel_df.to_excel(writer, sheet_name="Decibel Drop Data", index=False)
    
    # Melakukan ANOVA
    nrc_anova = analyze_anova(nrc_df, "NRC")
    decibel_anova = analyze_anova(decibel_df, "DecibelDrop")
    
    # Menyimpan hasil ANOVA
    with pd.ExcelWriter("anova_results.xlsx") as writer:
        nrc_anova.to_excel(writer, sheet_name="NRC ANOVA")
        decibel_anova.to_excel(writer, sheet_name="Decibel Drop ANOVA")
    
    # Membuat plot interaksi
    create_interaction_plots(nrc_df, "NRC")
    create_interaction_plots(decibel_df, "DecibelDrop")
    
    # Melakukan analisis LSD
    nrc_lsd = perform_lsd_analysis(nrc_df, "NRC")
    decibel_lsd = perform_lsd_analysis(decibel_df, "DecibelDrop")
    
    # Menyimpan hasil LSD
    with pd.ExcelWriter("lsd_results.xlsx") as writer:
        for factor, result in nrc_lsd.items():
            result.to_excel(writer, sheet_name=f"NRC {factor}", index=False)
        
        for factor, result in decibel_lsd.items():
            result.to_excel(writer, sheet_name=f"Decibel {factor}", index=False)
    
    print("Analisis selesai. Hasil disimpan dalam file Excel.") 