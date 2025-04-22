import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import statsmodels.api as sm
from statsmodels.formula.api import ols
from statsmodels.stats.anova import anova_lm
from statsmodels.stats.multicomp import pairwise_tukeyhsd
from statsmodels.stats.diagnostic import het_breuschpagan
from statsmodels.stats.stattools import durbin_watson
from scipy import stats
import glob
import os

# 1. Membaca data dari file Excel
def read_data_from_excel(file_path="input/semua_tabel.xlsx"):
    try:
        # Membaca data dari Excel
        nrc_sheet = "tabel_4.1"
        decibel_sheet = "tabel_4.2"
        
        # Periksa apakah file ada
        if not os.path.exists(file_path):
            print(f"Error: File {file_path} tidak ditemukan.")
            return None, None
        
        # Membaca data NRC (skiprows=1 karena baris pertama adalah judul)
        nrc_data = pd.read_excel(file_path, sheet_name=nrc_sheet, skiprows=1)
        
        # Membaca data Decibel Drop
        decibel_data = pd.read_excel(file_path, sheet_name=decibel_sheet, skiprows=1)
        
        # Parsing data ke dalam format yang dapat dianalisis
        compositions = ["50:50", "70:30", "90:10"]
        compactions = ["3:4", "4:4", "5:4"]
        cavities = ["15mm", "20mm", "25mm"]
        
        # Struktur data untuk menyimpan semua nilai
        nrc_df_data = []
        decibel_df_data = []
        
        # Untuk setiap kombinasi faktor
        for comp_idx, comp in enumerate(compositions):
            for compac_idx, compac in enumerate(compactions):
                for cav_idx, cav in enumerate(cavities):
                    # Menghitung indeks baris dalam Excel
                    row_idx = comp_idx * 9 + compac_idx * 3 + cav_idx
                    
                    try:
                        # Ekstrak nilai NRC (replikasi 1-5)
                        nrc_values = nrc_data.iloc[row_idx, 1:6].values
                        
                        # Ekstrak nilai Decibel Drop (replikasi 1-5)
                        decibel_values = decibel_data.iloc[row_idx, 1:6].values
                        
                        # Tambahkan data ke daftar
                        for rep in range(5):
                            nrc_df_data.append({
                                'Komposisi': comp,
                                'Kompaksi': compac,
                                'Cavity': cav,
                                'Replikasi': rep + 1,
                                'NRC': nrc_values[rep]
                            })
                            
                            decibel_df_data.append({
                                'Komposisi': comp,
                                'Kompaksi': compac,
                                'Cavity': cav,
                                'Replikasi': rep + 1,
                                'DecibelDrop': decibel_values[rep]
                            })
                    except Exception as e:
                        print(f"Error saat memproses data baris {row_idx}: {e}")
                        # Tambahkan data kosong jika error
                        for rep in range(5):
                            nrc_df_data.append({
                                'Komposisi': comp,
                                'Kompaksi': compac,
                                'Cavity': cav,
                                'Replikasi': rep + 1,
                                'NRC': np.nan
                            })
                            
                            decibel_df_data.append({
                                'Komposisi': comp,
                                'Kompaksi': compac,
                                'Cavity': cav,
                                'Replikasi': rep + 1,
                                'DecibelDrop': np.nan
                            })
        
        # Membuat DataFrame
        nrc_df = pd.DataFrame(nrc_df_data)
        decibel_df = pd.DataFrame(decibel_df_data)
        
        return nrc_df, decibel_df
        
    except Exception as e:
        print(f"Error saat membaca data dari Excel: {e}")
        return None, None

# 2. Melakukan analisis ANOVA dan menyimpan hasilnya
def analyze_anova(data, response_var, output_file=None):
    # Menentukan formula untuk ANOVA
    formula = f"{response_var} ~ C(Komposisi) + C(Kompaksi) + C(Cavity) + C(Komposisi):C(Kompaksi) + C(Komposisi):C(Cavity) + C(Kompaksi):C(Cavity) + C(Komposisi):C(Kompaksi):C(Cavity)"
    
    # Membuat model
    model = ols(formula, data=data).fit()
    
    # Melakukan ANOVA
    anova_result = anova_lm(model)
    
    # Menyimpan hasil jika diperlukan
    if output_file:
        # Memastikan direktori output ada
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        anova_result.to_csv(output_file)
        print(f"Hasil ANOVA disimpan di {output_file}")
    
    return model, anova_result

# 3. Melakukan post-hoc test (Tukey HSD)
def tukey_hsd_test(data, factor, response_var, alpha=0.05, output_file=None):
    # Melakukan Tukey HSD test
    tukey = pairwise_tukeyhsd(
        endog=data[response_var],
        groups=data[factor],
        alpha=alpha
    )
    
    # Konversi hasil ke DataFrame
    result_df = pd.DataFrame(
        data=tukey._results_table.data[1:], 
        columns=tukey._results_table.data[0]
    )
    
    # Menyimpan hasil jika diperlukan
    if output_file:
        # Memastikan direktori output ada
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        result_df.to_csv(output_file, index=False)
        print(f"Hasil Tukey HSD untuk {factor} disimpan di {output_file}")
    
    return tukey, result_df

# 4. Melakukan LSD (Least Significant Difference) test
def lsd_test(data, factor, response_var, model, alpha=0.05, output_file=None):
    # Group by factor dan hitung mean
    means = data.groupby(factor)[response_var].mean().reset_index()
    means = means.sort_values(response_var, ascending=False)
    
    # Hitung MSE dan df error dari hasil ANOVA
    anova_result = anova_lm(model)
    mse = anova_result.loc['Residual', 'mean_sq']
    df_error = anova_result.loc['Residual', 'df']
    
    # Hitung jumlah replikasi per level
    n_per_level = data.groupby(factor).size().iloc[0]
    
    # Hitung nilai t-kritis
    t_critical = stats.t.ppf(1 - alpha/2, df_error)
    
    # Hitung LSD
    lsd_value = t_critical * np.sqrt(2 * mse / n_per_level)
    
    # Membuat pasangan perbandingan dan menghitung perbedaan
    results = []
    for i, row1 in means.iterrows():
        for j, row2 in means.iterrows():
            if i < j:
                level1 = row1[factor]
                level2 = row2[factor]
                mean1 = row1[response_var]
                mean2 = row2[response_var]
                diff = abs(mean1 - mean2)
                significant = diff > lsd_value
                
                results.append({
                    'Factor': factor,
                    'Level 1': level1,
                    'Level 2': level2,
                    'Mean 1': mean1,
                    'Mean 2': mean2,
                    'Difference': diff,
                    'LSD Value': lsd_value,
                    'Significant': 'Yes' if significant else 'No'
                })
    
    # Konversi hasil ke DataFrame
    result_df = pd.DataFrame(results)
    
    # Menyimpan hasil jika diperlukan
    if output_file:
        # Memastikan direktori output ada
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        result_df.to_csv(output_file, index=False)
        print(f"Hasil LSD test untuk {factor} disimpan di {output_file}")
    
    return lsd_value, result_df

# 5. Uji asumsi normalitas
def test_normality(model, output_dir="output/plots"):
    # Memastikan direktori output ada
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Mendapatkan residual
    residuals = model.resid
    
    # 1. Uji Shapiro-Wilk
    shapiro_test = stats.shapiro(residuals)
    shapiro_pvalue = shapiro_test[1]
    shapiro_normal = shapiro_pvalue > 0.05
    
    # 2. QQ Plot
    plt.figure(figsize=(10, 6))
    fig = sm.qqplot(residuals, line='45', fit=True)
    plt.title('Normal Q-Q Plot of Residuals')
    plt.savefig(f"{output_dir}/normality_qqplot.png")
    plt.close()
    
    # 3. Histogram
    plt.figure(figsize=(10, 6))
    sns.histplot(residuals, kde=True)
    plt.title('Histogram of Residuals')
    plt.savefig(f"{output_dir}/normality_histogram.png")
    plt.close()
    
    return {
        'shapiro_test': shapiro_test,
        'is_normal': shapiro_normal,
        'p_value': shapiro_pvalue
    }

# 6. Uji asumsi homogenitas variansi
def test_homogeneity(model, data, factors, output_dir="output/plots"):
    # Memastikan direktori output ada
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Mendapatkan residual dan predicted values
    residuals = model.resid
    fitted = model.fittedvalues
    
    # 1. Breusch-Pagan test untuk heteroskedastisitas
    bp_test = het_breuschpagan(residuals, model.model.exog)
    bp_pvalue = bp_test[1]
    bp_homogen = bp_pvalue > 0.05
    
    # 2. Plot residual vs fitted values
    plt.figure(figsize=(10, 6))
    plt.scatter(fitted, residuals)
    plt.axhline(y=0, color='r', linestyle='-')
    plt.xlabel('Fitted values')
    plt.ylabel('Residuals')
    plt.title('Residuals vs Fitted Values')
    plt.savefig(f"{output_dir}/homogeneity_residual_vs_fitted.png")
    plt.close()
    
    # 3. Plot untuk setiap faktor
    for factor in factors:
        plt.figure(figsize=(10, 6))
        
        # Menghitung mean dan standard deviation untuk setiap level faktor
        stats_df = data.groupby(factor).agg({model.model.endog_names: ['mean', 'std']})
        stats_df.columns = stats_df.columns.droplevel()
        
        # Membuat bar plot untuk std
        ax = stats_df['std'].plot(kind='bar', color='skyblue')
        
        plt.title(f'Standard Deviation by {factor} Levels')
        plt.ylabel('Standard Deviation')
        plt.xlabel(factor)
        plt.tight_layout()
        plt.savefig(f"{output_dir}/homogeneity_{factor}_stdev.png")
        plt.close()
    
    return {
        'bp_test': bp_test,
        'is_homogeneous': bp_homogen,
        'p_value': bp_pvalue
    }

# 7. Uji asumsi independensi
def test_independence(model, data, output_dir="output/plots"):
    # Memastikan direktori output ada
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Mendapatkan residual
    residuals = model.resid
    
    # 1. Durbin-Watson test
    dw_test = durbin_watson(residuals)
    # DW test value range: 0-4, nilai sekitar 2 menunjukkan independensi
    dw_independent = 1.5 < dw_test < 2.5
    
    # 2. Plot residual vs order
    plt.figure(figsize=(10, 6))
    plt.plot(range(len(residuals)), residuals, 'o-')
    plt.axhline(y=0, color='r', linestyle='-')
    plt.title('Residuals vs Observation Order')
    plt.xlabel('Observation Order')
    plt.ylabel('Residuals')
    plt.savefig(f"{output_dir}/independence_residual_vs_order.png")
    plt.close()
    
    # 3. Autocorrelation plot
    plt.figure(figsize=(10, 6))
    pd.Series(residuals).autocorr_plot()
    plt.title('Autocorrelation Plot of Residuals')
    plt.savefig(f"{output_dir}/independence_autocorrelation.png")
    plt.close()
    
    return {
        'dw_test': dw_test,
        'is_independent': dw_independent
    }

# 8. Membuat plot interaksi
def create_interaction_plots(data, response_var, output_dir="output/plots"):
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
    plt.close()
    
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
    plt.close()
    
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
    plt.close()

# 9. Membuat main effects plot
def create_main_effects_plots(data, response_var, output_dir="output/plots"):
    # Memastikan direktori output ada
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    factors = ['Komposisi', 'Kompaksi', 'Cavity']
    
    # Membuat main effects plot untuk setiap faktor
    for factor in factors:
        # Menghitung mean untuk setiap level faktor
        means = data.groupby(factor)[response_var].mean()
        
        plt.figure(figsize=(10, 6))
        means.plot(kind='line', marker='o')
        plt.title(f'Main Effects Plot: {factor}')
        plt.xlabel(factor)
        plt.ylabel(f'Mean of {response_var}')
        plt.grid(True)
        plt.savefig(f"{output_dir}/main_effects_{factor}_{response_var}.png")
        plt.close()

# 10. Membuat surface plot untuk visualisasi 3D
def create_surface_plots(data, response_var, output_dir="output/plots"):
    # Memastikan direktori output ada
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Menghitung means untuk pasangan faktor
    factor_pairs = [('Komposisi', 'Kompaksi'), ('Komposisi', 'Cavity'), ('Kompaksi', 'Cavity')]
    
    for factor1, factor2 in factor_pairs:
        # Aggregasi data
        pivot_data = data.groupby([factor1, factor2])[response_var].mean().reset_index()
        pivot_table = pivot_data.pivot(index=factor1, columns=factor2, values=response_var)
        
        # Membuat meshgrid untuk surface plot
        x = np.arange(len(pivot_table.index))
        y = np.arange(len(pivot_table.columns))
        X, Y = np.meshgrid(x, y)
        Z = pivot_table.values.T
        
        # Membuat 3D surface plot
        fig = plt.figure(figsize=(12, 8))
        ax = fig.add_subplot(111, projection='3d')
        surf = ax.plot_surface(X, Y, Z, cmap='viridis', alpha=0.8)
        
        # Mengatur label dan judul
        ax.set_xlabel(factor1)
        ax.set_ylabel(factor2)
        ax.set_zlabel(response_var)
        ax.set_title(f'Surface Plot: {factor1} vs {factor2} untuk {response_var}')
        
        # Mengatur ticks
        ax.set_xticks(x)
        ax.set_xticklabels(pivot_table.index)
        ax.set_yticks(y)
        ax.set_yticklabels(pivot_table.columns)
        
        # Menambahkan colorbar
        fig.colorbar(surf, ax=ax, shrink=0.5, aspect=5)
        
        # Menyimpan plot
        plt.savefig(f"{output_dir}/surface_plot_{factor1}_{factor2}_{response_var}.png")
        plt.close()

# 11. Menyimpan hasil analisis ke Excel
def save_analysis_summary(nrc_model, decibel_model, nrc_anova, decibel_anova, output_file="output/analysis_summary.xlsx"):
    # Memastikan direktori output ada
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    
    # Menyimpan ke Excel
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Menyimpan hasil ANOVA
        nrc_anova.to_excel(writer, sheet_name="NRC ANOVA")
        decibel_anova.to_excel(writer, sheet_name="Decibel ANOVA")
        
        # Menyimpan parameter model
        pd.DataFrame(nrc_model.params).to_excel(writer, sheet_name="NRC Parameters")
        pd.DataFrame(decibel_model.params).to_excel(writer, sheet_name="Decibel Parameters")
        
        # Menyimpan statistik model
        model_stats_nrc = {
            'R-squared': nrc_model.rsquared,
            'Adj. R-squared': nrc_model.rsquared_adj,
            'F-statistic': nrc_model.fvalue,
            'Prob (F-statistic)': nrc_model.f_pvalue,
            'AIC': nrc_model.aic,
            'BIC': nrc_model.bic
        }
        
        model_stats_decibel = {
            'R-squared': decibel_model.rsquared,
            'Adj. R-squared': decibel_model.rsquared_adj,
            'F-statistic': decibel_model.fvalue,
            'Prob (F-statistic)': decibel_model.f_pvalue,
            'AIC': decibel_model.aic,
            'BIC': decibel_model.bic
        }
        
        pd.DataFrame([model_stats_nrc]).to_excel(writer, sheet_name="NRC Model Stats")
        pd.DataFrame([model_stats_decibel]).to_excel(writer, sheet_name="Decibel Model Stats")
    
    print(f"Ringkasan analisis telah disimpan di {output_file}")

# Menjalankan semua analisis
if __name__ == "__main__":
    # 1. Membaca data dari Excel
    print("Membaca data dari file Excel...")
    nrc_df, decibel_df = read_data_from_excel()
    
    if nrc_df is None or decibel_df is None:
        print("Error: Tidak dapat membaca data dari Excel. Pastikan file semua_tabel.xlsx ada di folder 'input'.")
        exit()
    
    # Membuat direktori untuk output
    os.makedirs("output/results", exist_ok=True)
    os.makedirs("output/plots/nrc", exist_ok=True)
    os.makedirs("output/plots/decibel", exist_ok=True)
    
    # 2. Analisis untuk data NRC
    print("\nMelakukan analisis untuk data NRC...")
    nrc_model, nrc_anova = analyze_anova(nrc_df, "NRC", "output/results/nrc_anova.csv")
    
    # 3. Analisis untuk data Decibel Drop
    print("\nMelakukan analisis untuk data DecibelDrop...")
    decibel_model, decibel_anova = analyze_anova(decibel_df, "DecibelDrop", "output/results/decibel_anova.csv")
    
    # 4. Post-hoc tests untuk data NRC
    print("\nMelakukan post-hoc tests untuk data NRC...")
    for factor in ['Komposisi', 'Kompaksi', 'Cavity']:
        # Tukey HSD
        tukey_hsd_test(nrc_df, factor, "NRC", output_file=f"output/results/nrc_tukey_{factor}.csv")
        
        # LSD test
        lsd_test(nrc_df, factor, "NRC", nrc_model, output_file=f"output/results/nrc_lsd_{factor}.csv")
    
    # 5. Post-hoc tests untuk data Decibel Drop
    print("\nMelakukan post-hoc tests untuk data DecibelDrop...")
    for factor in ['Komposisi', 'Kompaksi', 'Cavity']:
        # Tukey HSD
        tukey_hsd_test(decibel_df, factor, "DecibelDrop", output_file=f"output/results/decibel_tukey_{factor}.csv")
        
        # LSD test
        lsd_test(decibel_df, factor, "DecibelDrop", decibel_model, output_file=f"output/results/decibel_lsd_{factor}.csv")
    
    # 6. Uji asumsi untuk NRC
    print("\nMelakukan uji asumsi ANOVA untuk data NRC...")
    # Normalitas
    norm_result_nrc = test_normality(nrc_model, "output/plots/nrc")
    print(f"Normalitas (NRC): p-value = {norm_result_nrc['p_value']:.4f}, Normal: {norm_result_nrc['is_normal']}")
    
    # Homogenitas variansi
    homo_result_nrc = test_homogeneity(nrc_model, nrc_df, ['Komposisi', 'Kompaksi', 'Cavity'], "output/plots/nrc")
    print(f"Homogenitas (NRC): p-value = {homo_result_nrc['p_value']:.4f}, Homogen: {homo_result_nrc['is_homogeneous']}")
    
    # Independensi
    indep_result_nrc = test_independence(nrc_model, nrc_df, "output/plots/nrc")
    print(f"Independensi (NRC): DW = {indep_result_nrc['dw_test']:.4f}, Independent: {indep_result_nrc['is_independent']}")
    
    # 7. Uji asumsi untuk Decibel Drop
    print("\nMelakukan uji asumsi ANOVA untuk data DecibelDrop...")
    # Normalitas
    norm_result_decibel = test_normality(decibel_model, "output/plots/decibel")
    print(f"Normalitas (DecibelDrop): p-value = {norm_result_decibel['p_value']:.4f}, Normal: {norm_result_decibel['is_normal']}")
    
    # Homogenitas variansi
    homo_result_decibel = test_homogeneity(decibel_model, decibel_df, ['Komposisi', 'Kompaksi', 'Cavity'], "output/plots/decibel")
    print(f"Homogenitas (DecibelDrop): p-value = {homo_result_decibel['p_value']:.4f}, Homogen: {homo_result_decibel['is_homogeneous']}")
    
    # Independensi
    indep_result_decibel = test_independence(decibel_model, decibel_df, "output/plots/decibel")
    print(f"Independensi (DecibelDrop): DW = {indep_result_decibel['dw_test']:.4f}, Independent: {indep_result_decibel['is_independent']}")
    
    # 8. Membuat plot interaksi
    print("\nMembuat plot interaksi...")
    create_interaction_plots(nrc_df, "NRC", "output/plots/nrc")
    create_interaction_plots(decibel_df, "DecibelDrop", "output/plots/decibel")
    
    # 9. Membuat main effects plots
    print("\nMembuat main effects plots...")
    create_main_effects_plots(nrc_df, "NRC", "output/plots/nrc")
    create_main_effects_plots(decibel_df, "DecibelDrop", "output/plots/decibel")
    
    # 10. Membuat surface plots
    print("\nMembuat surface plots...")
    create_surface_plots(nrc_df, "NRC", "output/plots/nrc")
    create_surface_plots(decibel_df, "DecibelDrop", "output/plots/decibel")
    
    # 11. Menyimpan ringkasan analisis
    print("\nMenyimpan ringkasan analisis...")
    save_analysis_summary(nrc_model, decibel_model, nrc_anova, decibel_anova)
    
    print("\nAnalisis selesai! Hasil tersimpan di folder 'output'.") 