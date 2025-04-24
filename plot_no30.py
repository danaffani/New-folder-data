import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import scipy.stats as stats
from statsmodels.stats.diagnostic import het_breuschpagan
from statsmodels.stats.stattools import durbin_watson
from statsmodels.graphics.gofplots import qqplot
from statsmodels.formula.api import ols
import statsmodels.api as sm
import os
from statsmodels.stats.outliers_influence import OLSInfluence
from statsmodels.stats.multicomp import pairwise_tukeyhsd

os.makedirs('output/plot', exist_ok=True)

def plot_homogeneity(data, group_col, value_col):
    """
    Plot untuk uji homogenitas varians
    
    Parameters:
    -----------
    data : DataFrame
        Data dalam format pandas DataFrame
    group_col : str
        Nama kolom untuk pengelompokan
    value_col : str
        Nama kolom untuk nilai yang dianalisis
    """
    plt.figure(figsize=(10, 6))
    ax1 = plt.subplot(121)
    sns.boxplot(x=group_col, y=value_col, data=data, ax=ax1)
    ax1.set_title('Boxplot untuk Cek Homogenitas Varians')
    
    model = ols(f'{value_col} ~ C({group_col})', data=data).fit()
    residuals = model.resid
    fitted_values = model.fittedvalues
    
    ax2 = plt.subplot(122)
    sns.scatterplot(x=fitted_values, y=residuals, ax=ax2)
    ax2.axhline(y=0, color='r', linestyle='-')
    ax2.set_title('Residual Plot untuk Homoskedastisitas')
    ax2.set_xlabel('Fitted Values')
    ax2.set_ylabel('Residuals')
    
    plt.tight_layout()
    plt.savefig('output/plot/homogeneity_plot.png')
    plt.close()
    
    groups = [data[data[group_col] == group][value_col].values for group in data[group_col].unique()]
    stat, p_value = stats.levene(*groups)
    print(f"Levene's Test: Statistic={stat:.4f}, p-value={p_value:.4f}")
    if p_value > 0.05:
        print("Asumsi homogenitas varians terpenuhi (p > 0.05)")
    else:
        print("Asumsi homogenitas varians tidak terpenuhi (p <= 0.05)")
    
    bp_test = het_breuschpagan(residuals, model.model.exog)
    _, bp_p_value, _, _ = bp_test
    print(f"Breusch-Pagan Test: p-value={bp_p_value:.4f}")
    if bp_p_value > 0.05:
        print("Asumsi homoskedastisitas terpenuhi (p > 0.05)")
    else:
        print("Asumsi homoskedastisitas tidak terpenuhi (p <= 0.05)")

def plot_normality(data, group_col, value_col):
    """
    Plot untuk uji normalitas residual
    
    Parameters:
    -----------
    data : DataFrame
        Data dalam format pandas DataFrame
    group_col : str
        Nama kolom untuk pengelompokan
    value_col : str
        Nama kolom untuk nilai yang dianalisis
    """
    model = ols(f'{value_col} ~ C({group_col})', data=data).fit()
    residuals = model.resid
    
    plt.figure(figsize=(12, 6))
    
    ax1 = plt.subplot(121)
    sns.histplot(residuals, kde=True, ax=ax1)
    ax1.set_title('Histogram Residual')
    ax1.set_xlabel('Residual')
    
    ax2 = plt.subplot(122)
    qqplot(residuals, line='s', ax=ax2)
    ax2.set_title('Q-Q Plot Residual')
    
    plt.tight_layout()
    plt.savefig('output/plot/normality_plot.png')
    plt.close()
    
    stat, p_value = stats.shapiro(residuals)
    print(f"Shapiro-Wilk Test: Statistic={stat:.4f}, p-value={p_value:.4f}")
    if p_value > 0.05:
        print("Asumsi normalitas terpenuhi (p > 0.05)")
    else:
        print("Asumsi normalitas tidak terpenuhi (p <= 0.05)")
    
    print("\nUji Normalitas per Grup:")
    for group in data[group_col].unique():
        group_data = data[data[group_col] == group][value_col]
        stat, p_value = stats.shapiro(group_data)
        print(f"Grup {group} - Shapiro-Wilk Test: Statistic={stat:.4f}, p-value={p_value:.4f}")

def plot_independence(data, group_col, value_col):
    """
    Plot untuk uji independensi residual
    
    Parameters:
    -----------
    data : DataFrame
        Data dalam format pandas DataFrame
    group_col : str
        Nama kolom untuk pengelompokan
    value_col : str
        Nama kolom untuk nilai yang dianalisis
    """
    model = ols(f'{value_col} ~ C({group_col})', data=data).fit()
    residuals = model.resid
    
    plt.figure(figsize=(10, 6))
    
    plt.scatter(range(len(residuals)), residuals)
    plt.axhline(y=0, color='r', linestyle='-')
    plt.title('Residual Plot untuk Independensi')
    plt.xlabel('Urutan Observasi')
    plt.ylabel('Residual')
    
    plt.tight_layout()
    plt.savefig('output/plot/independence_plot.png')
    plt.close()
    
    dw_stat = durbin_watson(residuals)
    print(f"Durbin-Watson Statistic: {dw_stat:.4f}")
    if dw_stat < 1.5:
        print("Nilai DW < 1.5: Kemungkinan ada autokorelasi positif")
    elif dw_stat > 2.5:
        print("Nilai DW > 2.5: Kemungkinan ada autokorelasi negatif")
    else:
        print("Nilai DW sekitar 2: Tidak ada autokorelasi")

def plot_outliers(data, group_col, value_col):
    """
    Plot untuk deteksi outlier
    
    Parameters:
    -----------
    data : DataFrame
        Data dalam format pandas DataFrame
    group_col : str
        Nama kolom untuk pengelompokan
    value_col : str
        Nama kolom untuk nilai yang dianalisis
    """
    plt.figure(figsize=(12, 10))
    
    ax1 = plt.subplot(221)
    sns.boxplot(x=group_col, y=value_col, data=data, ax=ax1)
    sns.swarmplot(x=group_col, y=value_col, data=data, color='black', alpha=0.5, ax=ax1)
    ax1.set_title('Boxplot dengan Data Points')
    
    model = ols(f'{value_col} ~ C({group_col})', data=data).fit()
    influence = OLSInfluence(model)
    leverage = influence.hat_matrix_diag
    
    ax2 = plt.subplot(222)
    ax2.scatter(range(len(leverage)), leverage)
    ax2.set_title('Leverage Values')
    ax2.set_xlabel('Observation Index')
    ax2.set_ylabel('Leverage')
    threshold = 2 * (model.df_model + 1) / len(data)
    ax2.axhline(y=threshold, color='r', linestyle='--', label=f'Threshold ({threshold:.3f})')
    ax2.legend()
    
    cooks_d = influence.cooks_distance[0]
    
    ax3 = plt.subplot(223)
    ax3.stem(range(len(cooks_d)), cooks_d, markerfmt=",")
    ax3.set_title("Cook's Distance")
    ax3.set_xlabel('Observation Index')
    ax3.set_ylabel("Cook's Distance")
    
    ax4 = plt.subplot(224)
    data['abs_resid'] = np.abs(model.resid)
    sns.boxplot(x=group_col, y='abs_resid', data=data, ax=ax4)
    ax4.set_title('Absolute Residuals per Group')
    ax4.set_ylabel('Absolute Residual')
    
    plt.tight_layout()
    plt.savefig('output/plot/outliers_plot.png')
    plt.close()
    
    threshold_cooks = 4 / len(data)
    outliers_idx = np.where(cooks_d > threshold_cooks)[0]
    
    if len(outliers_idx) > 0:
        print(f"Potensial outlier terdeteksi pada indeks: {outliers_idx}")
        print("Data pada indeks tersebut:")
        for idx in outliers_idx:
            group = data.iloc[idx][group_col]
            value = data.iloc[idx][value_col]
            print(f"Indeks {idx}: {group_col}={group}, {value_col}={value:.3f}, Cook's D={cooks_d[idx]:.3f}")
    else:
        print("Tidak ada outlier signifikan yang terdeteksi")

def plot_qq_per_group(data, group_col, value_col):
    """
    Plot QQ normal untuk setiap grup
    
    Parameters:
    -----------
    data : DataFrame
        Data dalam format pandas DataFrame
    group_col : str
        Nama kolom untuk pengelompokan
    value_col : str
        Nama kolom untuk nilai yang dianalisis
    """
    groups = data[group_col].unique()
    n_groups = len(groups)
    
    if n_groups <= 3:
        n_rows, n_cols = 1, n_groups
    else:
        n_rows = (n_groups + 1) // 2
        n_cols = 2
    
    plt.figure(figsize=(n_cols * 6, n_rows * 5))
    
    for i, group in enumerate(groups):
        ax = plt.subplot(n_rows, n_cols, i+1)
        group_data = data[data[group_col] == group][value_col]
        qqplot(group_data, line='s', ax=ax)
        ax.set_title(f'Q-Q Plot: {group}')
    
    plt.tight_layout()
    plt.savefig('output/plot/qq_per_group_plot.png')
    plt.close()

def plot_means(data, group_col, value_col):
    """
    Plot rata-rata dengan interval kepercayaan 95%
    
    Parameters:
    -----------
    data : DataFrame
        Data dalam format pandas DataFrame
    group_col : str
        Nama kolom untuk pengelompokan
    value_col : str
        Nama kolom untuk nilai yang dianalisis
    """
    plt.figure(figsize=(10, 6))
    
    ax1 = plt.subplot(121)
    sns.barplot(x=group_col, y=value_col, data=data, errorbar=('ci', 95), ax=ax1)
    ax1.set_title('Mean dengan 95% CI')
    ax1.set_ylabel(value_col)
    
    ax2 = plt.subplot(122)
    sns.pointplot(x=group_col, y=value_col, data=data, errorbar=('ci', 95), capsize=0.1, ax=ax2)
    ax2.set_title('Point Plot dengan 95% CI')
    ax2.set_ylabel(value_col)
    
    plt.tight_layout()
    plt.savefig('output/plot/means_plot.png')
    plt.close()
    
    print("Statistik Deskriptif per Grup:")
    stats_df = data.groupby(group_col)[value_col].agg(['count', 'mean', 'std', 'min', 'max'])
    stats_df['sem'] = data.groupby(group_col)[value_col].sem()
    stats_df['ci95'] = stats_df['sem'] * 1.96  #95% CI approxx
    print(stats_df)

def plot_posthoc(data, group_col, value_col):
    """
    Plot hasil post-hoc test (Tukey HSD)
    
    Parameters:
    -----------
    data : DataFrame
        Data dalam format pandas DataFrame
    group_col : str
        Nama kolom untuk pengelompokan
    value_col : str
        Nama kolom untuk nilai yang dianalisis
    """
    model = ols(f'{value_col} ~ C({group_col})', data=data).fit()
    anova_table = sm.stats.anova_lm(model, typ=2)
    print("ANOVA Results:")
    print(anova_table)
    
    tukey = pairwise_tukeyhsd(endog=data[value_col], groups=data[group_col], alpha=0.05)
    print("\nTukey HSD Results:")
    print(tukey)
    
    plt.figure(figsize=(10, 6))
    
    ax1 = plt.subplot(121)
    
    groups = tukey.groupsunique
    n_groups = len(groups)
    mean_diffs = []
    lower_bounds = []
    upper_bounds = []
    p_values = []
    labels = []
    
    for i in range(n_groups):
        for j in range(i+1, n_groups):
            idx = None
            for k, row in enumerate(tukey._results_table.data):
                if row[0] == groups[i] and row[1] == groups[j]:
                    idx = k
                    break
            
            if idx is not None:
                mean_diff = tukey.meandiffs[idx-1]  #minus 1 karna baris pertama adalah header
                lower = tukey.confint[idx-1, 0]
                upper = tukey.confint[idx-1, 1]
                p_value = tukey.pvalues[idx-1]
                
                mean_diffs.append(mean_diff)
                lower_bounds.append(lower)
                upper_bounds.append(upper)
                p_values.append(p_value)
                labels.append(f"{groups[i]} - {groups[j]}")
    
    positions = range(len(mean_diffs))
    ax1.errorbar(
        x=mean_diffs,
        y=positions,
        xerr=np.vstack([(np.array(mean_diffs) - np.array(lower_bounds)), 
                         (np.array(upper_bounds) - np.array(mean_diffs))]),
        fmt='o', capsize=5, elinewidth=2, markeredgewidth=2
    )
    ax1.axvline(x=0, color='r', linestyle='--')
    ax1.set_yticks(positions)
    ax1.set_yticklabels(labels)
    ax1.set_xlabel('Mean Difference')
    ax1.set_title('Tukey HSD: Mean Differences dengan 95% CI')
    
    p_matrix = np.ones((n_groups, n_groups))
    
    count = 0
    for i in range(n_groups):
        for j in range(i+1, n_groups):
            p_matrix[i, j] = p_values[count]
            p_matrix[j, i] = p_values[count]  # matrix simetris
            count += 1
    
    ax2 = plt.subplot(122)
    im = ax2.imshow(p_matrix, cmap='YlOrRd_r', vmin=0, vmax=0.1)
    
    # Annotasi p-values
    for i in range(n_groups):
        for j in range(n_groups):
            if i != j:
                if p_matrix[i, j] < 0.001:
                    text = f"{p_matrix[i, j]:.2e}"
                else:
                    text = f"{p_matrix[i, j]:.3f}"
                color = 'white' if p_matrix[i, j] < 0.05 else 'black'
                ax2.text(j, i, text, ha='center', va='center', color=color)
            else:
                ax2.text(j, i, '-', ha='center', va='center')
    
    plt.colorbar(im, label='p-value')
    ax2.set_xticks(range(n_groups))
    ax2.set_yticks(range(n_groups))
    ax2.set_xticklabels(groups)
    ax2.set_yticklabels(groups)
    ax2.set_title('Tukey HSD: p-values')
    
    plt.tight_layout()
    plt.savefig('output/plot/posthoc_plot.png')
    plt.close()

def plot_interaction(data, factor1, factor2, value_col):
    """
    Plot interaksi untuk ANOVA faktorial
    
    Parameters:
    -----------
    data : DataFrame
        Data dalam format pandas DataFrame
    factor1 : str
        Nama kolom untuk faktor pertama
    factor2 : str
        Nama kolom untuk faktor kedua
    value_col : str
        Nama kolom untuk nilai yang dianalisis
    """
    if factor1 not in data.columns or factor2 not in data.columns:
        print(f"Error: Faktor '{factor1}' atau '{factor2}' tidak ditemukan di data")
        return
    
    plt.figure(figsize=(15, 6))
    
    ax1 = plt.subplot(121)
    sns.pointplot(x=factor1, y=value_col, hue=factor2, data=data, errorbar=('ci', 95), ax=ax1)
    ax1.set_title(f'Interaction Plot: {factor1} x {factor2}')
    
    ax2 = plt.subplot(122)
    sns.pointplot(x=factor2, y=value_col, hue=factor1, data=data, errorbar=('ci', 95), ax=ax2)
    ax2.set_title(f'Interaction Plot: {factor2} x {factor1}')
    
    plt.tight_layout()
    plt.savefig('output/plot/interaction_plot.png')
    plt.close()
    
    model = ols(f'{value_col} ~ C({factor1}) * C({factor2})', data=data).fit()
    anova_table = sm.stats.anova_lm(model, typ=2)
    print("Factorial ANOVA Results:")
    print(anova_table)
    
    interaction_p = anova_table.loc[f'C({factor1}):C({factor2})', 'PR(>F)']
    if interaction_p < 0.05:
        print(f"Interaksi antara {factor1} dan {factor2} signifikan (p = {interaction_p:.4f})")
    else:
        print(f"Interaksi antara {factor1} dan {factor2} tidak signifikan (p = {interaction_p:.4f})")

def plot_all_assumptions(data, group_col, value_col):
    """
    Plot semua asumsi ANOVA dan visualisasi hasil
    
    Parameters:
    -----------
    data : DataFrame
        Data dalam format pandas DataFrame
    group_col : str
        Nama kolom untuk pengelompokan
    value_col : str
        Nama kolom untuk nilai yang dianalisis
    """
    print("="*50)
    print("UJI ASUMSI HOMOGENITAS VARIANS")
    print("="*50)
    plot_homogeneity(data, group_col, value_col)
    
    print("\n"+"="*50)
    print("UJI ASUMSI NORMALITAS")
    print("="*50)
    plot_normality(data, group_col, value_col)
    
    print("\n"+"="*50)
    print("UJI ASUMSI INDEPENDENSI")
    print("="*50)
    plot_independence(data, group_col, value_col)
    
    print("\n"+"="*50)
    print("QQ PLOT PER GRUP")
    print("="*50)
    plot_qq_per_group(data, group_col, value_col)
    
    print("\n"+"="*50)
    print("PLOT RATA-RATA")
    print("="*50)
    plot_means(data, group_col, value_col)
    
    print("\n"+"="*50)
    print("ANOVA DAN POST-HOC TEST")
    print("="*50)
    plot_posthoc(data, group_col, value_col)

def contoh_penggunaan():
    """
    Contoh penggunaan fungsi dengan data simulasi
    """
    np.random.seed(42)
    group_sizes = [30, 30, 30]
    means = [5, 7, 9]
    
    data = []
    for i, (size, mean) in enumerate(zip(group_sizes, means)):
        group_data = np.random.normal(mean, 1.5, size)
        for val in group_data:
            data.append({
                'group': f'Group {i+1}',
                'value': val
            })
    
    df = pd.DataFrame(data)
    
    plot_all_assumptions(df, 'group', 'value')
    
    n_per_cell = 20
    factors = {'A': ['A1', 'A2'], 'B': ['B1', 'B2', 'B3']}
    
    factorial_data = []
    
    for level_a in factors['A']:
        for level_b in factors['B']:
            if level_a == 'A1':
                mean_a = 10
            else:
                mean_a = 15
                
            if level_b == 'B1':
                mean_b = 0
            elif level_b == 'B2':
                mean_b = 2
            else:
                mean_b = 4
                
            interaction = 0
            if level_a == 'A2' and level_b == 'B3':
                interaction = 3
                
            cell_mean = mean_a + mean_b + interaction
            cell_values = np.random.normal(cell_mean, 2, n_per_cell)
            
            for value in cell_values:
                factorial_data.append({
                    'factor_a': level_a,
                    'factor_b': level_b,
                    'response': value
                })
    
    factorial_df = pd.DataFrame(factorial_data)
    
    print("\n"+"="*50)
    print("PLOT INTERAKSI (ANOVA FAKTORIAL)")
    print("="*50)
    plot_interaction(factorial_df, 'factor_a', 'factor_b', 'response')

if __name__ == "__main__":
    try:
        anova_data = pd.read_excel('input/nomor8_tambahan.xlsx', sheet_name='Hasil ANOVA')
        print("Data ANOVA berhasil dimuat.")
        
        experiment_data = pd.read_excel('input/semua_tabel.xlsx', sheet_name='tabel_4.1')
        print("Data eksperimen berhasil dimuat.")

        df_long = pd.melt(experiment_data, 
                          id_vars=['Komposisi', 'Kompaksi', 'Cavity (mm)'],
                          value_vars=['NRC1', 'NRC2', 'NRC3', 'NRC4', 'NRC5'],
                          var_name='Replikasi', 
                          value_name='NRC')
        
        print(f"Bentuk data yang dianalisis: {df_long.shape}")
        print("Preview data:")
        print(df_long.head())
        
        plot_all_assumptions(df_long, 'Komposisi', 'NRC')
        
        print("\n"+"="*50)
        print("PLOT INTERAKSI (ANOVA FAKTORIAL)")
        print("="*50)
        plot_interaction(df_long, 'Komposisi', 'Kompaksi', 'NRC')
        
    except Exception as e:
        print(f"Error saat memuat atau memproses data: {e}")
        print("Menggunakan data simulasi sebagai contoh")
        contoh_penggunaan() 