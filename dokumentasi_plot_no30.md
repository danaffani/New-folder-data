# Dokumentasi plot_no30.py

## Tujuan Program

Program `plot_no30.py` bertujuan untuk menghasilkan visualisasi grafik yang dibutuhkan dalam analisis ANOVA (Analysis of Variance). Program ini dirancang khusus untuk:

1. Menguji asumsi-asumsi ANOVA secara visual dan statistik
2. Menghasilkan plot untuk interpretasi hasil ANOVA
3. Menampilkan hasil uji post-hoc dan interaksi antar faktor

## Library yang Digunakan

Program ini menggunakan beberapa library Python untuk analisis statistik dan visualisasi:

- **numpy**: Untuk operasi numerik dan manipulasi array
- **pandas**: Untuk pengelolaan dan manipulasi data dalam format DataFrame
- **matplotlib**: Untuk pembuatan visualisasi grafik dasar
- **seaborn**: Untuk visualisasi statistik yang lebih baik dan estetis
- **scipy.stats**: Untuk uji statistik (Levene, Shapiro-Wilk, dll)
- **statsmodels**: Untuk model statistik seperti ANOVA, uji diagnostik (Breusch-Pagan, Durbin-Watson) dan visualisasi khusus (QQ-plot)

## Struktur dan Fungsi Utama

Program terdiri dari beberapa fungsi utama yang masing-masing memiliki peran spesifik:

### 1. `plot_homogeneity(data, group_col, value_col)`

Fungsi ini menghasilkan plot dan uji statistik untuk mengevaluasi asumsi homogenitas varians (homoskedastisitas).

- **Contoh kode implementasi uji Levene**:
  ```python
  # Mengelompokkan data berdasarkan grup
  groups = [data[data[group_col] == group][value_col].values 
            for group in data[group_col].unique()]
  
  # Melakukan uji Levene
  stat, p_value = stats.levene(*groups)
  
  # Interpretasi hasil
  if p_value > 0.05:
      print("Asumsi homogenitas varians terpenuhi (p > 0.05)")
  else:
      print("Asumsi homogenitas varians tidak terpenuhi (p <= 0.05)")
  ```

- **Contoh kode implementasi uji Breusch-Pagan**:
  ```python
  # Menggunakan model regresi linear
  model = ols(f'{value_col} ~ C({group_col})', data=data).fit()
  residuals = model.resid
  
  # Melakukan uji Breusch-Pagan
  bp_test = het_breuschpagan(residuals, model.model.exog)
  _, bp_p_value, _, _ = bp_test
  
  # Interpretasi hasil
  if bp_p_value > 0.05:
      print("Asumsi homoskedastisitas terpenuhi (p > 0.05)")
  else:
      print("Asumsi homoskedastisitas tidak terpenuhi (p <= 0.05)")
  ```

### 2. `plot_normality(data, group_col, value_col)`

Fungsi untuk mengevaluasi asumsi normalitas residual.

- **Contoh kode implementasi uji Shapiro-Wilk**:
  ```python
  # Membuat model ANOVA
  model = ols(f'{value_col} ~ C({group_col})', data=data).fit()
  residuals = model.resid
  
  # Melakukan uji Shapiro-Wilk pada residual
  stat, p_value = stats.shapiro(residuals)
  
  # Interpretasi hasil
  if p_value > 0.05:
      print("Asumsi normalitas terpenuhi (p > 0.05)")
  else:
      print("Asumsi normalitas tidak terpenuhi (p <= 0.05)")
  
  # Juga melakukan uji per grup
  for group in data[group_col].unique():
      group_data = data[data[group_col] == group][value_col]
      stat, p_value = stats.shapiro(group_data)
      print(f"Grup {group} - Shapiro-Wilk Test: Statistic={stat:.4f}, p-value={p_value:.4f}")
  ```

### 3. `plot_independence(data, group_col, value_col)`

Fungsi untuk mengevaluasi asumsi independensi residual (tidak ada autokorelasi).

- **Contoh kode implementasi uji Durbin-Watson**:
  ```python
  # Membuat model ANOVA
  model = ols(f'{value_col} ~ C({group_col})', data=data).fit()
  residuals = model.resid
  
  # Melakukan uji Durbin-Watson
  dw_stat = durbin_watson(residuals)
  
  # Interpretasi hasil
  if dw_stat < 1.5:
      print("Nilai DW < 1.5: Kemungkinan ada autokorelasi positif")
  elif dw_stat > 2.5:
      print("Nilai DW > 2.5: Kemungkinan ada autokorelasi negatif")
  else:
      print("Nilai DW sekitar 2: Tidak ada autokorelasi")
  ```

### 4. `plot_qq_per_group(data, group_col, value_col)`

Menyediakan QQ-plot untuk setiap kelompok secara terpisah.

- **Contoh kode implementasi**:
  ```python
  # Mendapatkan semua grup unik
  groups = data[group_col].unique()
  n_groups = len(groups)
  
  # Menentukan layout plot
  if n_groups <= 3:
      n_rows, n_cols = 1, n_groups
  else:
      n_rows = (n_groups + 1) // 2
      n_cols = 2
  
  # Membuat QQ plot untuk setiap grup
  for i, group in enumerate(groups):
      ax = plt.subplot(n_rows, n_cols, i+1)
      group_data = data[data[group_col] == group][value_col]
      qqplot(group_data, line='s', ax=ax)
      ax.set_title(f'Q-Q Plot: {group}')
  ```

### 5. `plot_means(data, group_col, value_col)`

Visualisasi deskriptif dari rata-rata setiap kelompok.

- **Contoh kode untuk statistik deskriptif**:
  ```python
  # Menghitung statistik deskriptif
  stats_df = data.groupby(group_col)[value_col].agg(['count', 'mean', 'std', 'min', 'max'])
  
  # Menambahkan standard error of mean (SEM)
  stats_df['sem'] = data.groupby(group_col)[value_col].sem()
  
  # Menghitung 95% confidence interval
  stats_df['ci95'] = stats_df['sem'] * 1.96  # Aproksimasi CI 95%
  
  print("Statistik Deskriptif per Grup:")
  print(stats_df)
  ```

### 6. `plot_posthoc(data, group_col, value_col)`

Mengimplementasikan dan memvisualisasikan hasil uji post-hoc (Tukey HSD) setelah ANOVA.

- **Contoh kode implementasi ANOVA dan Tukey HSD**:
  ```python
  # Membuat model ANOVA
  model = ols(f'{value_col} ~ C({group_col})', data=data).fit()
  anova_table = sm.stats.anova_lm(model, typ=2)
  
  # Melakukan uji Tukey HSD
  tukey = pairwise_tukeyhsd(endog=data[value_col], groups=data[group_col], alpha=0.05)
  
  # Ekstraksi data untuk visualisasi
  groups = tukey.groupsunique
  n_groups = len(groups)
  mean_diffs = []
  lower_bounds = []
  upper_bounds = []
  p_values = []
  labels = []
  
  # Mendapatkan data perbandingan berpasangan
  for i in range(n_groups):
      for j in range(i+1, n_groups):
          idx = None
          for k, row in enumerate(tukey._results_table.data):
              if row[0] == groups[i] and row[1] == groups[j]:
                  idx = k
                  break
          
          if idx is not None:
              mean_diff = tukey.meandiffs[idx-1]  # Minus 1 karena baris pertama adalah header
              lower = tukey.confint[idx-1, 0]
              upper = tukey.confint[idx-1, 1]
              p_value = tukey.pvalues[idx-1]
              
              mean_diffs.append(mean_diff)
              lower_bounds.append(lower)
              upper_bounds.append(upper)
              p_values.append(p_value)
              labels.append(f"{groups[i]} - {groups[j]}")
  ```

- **Contoh kode visualisasi p-values**:
  ```python
  # Membuat matriks p-values untuk heatmap
  p_matrix = np.ones((n_groups, n_groups))
  
  count = 0
  for i in range(n_groups):
      for j in range(i+1, n_groups):
          p_matrix[i, j] = p_values[count]
          p_matrix[j, i] = p_values[count]  # Matrix simetris
          count += 1
  
  # Visualisasi p-values dengan heatmap
  im = ax2.imshow(p_matrix, cmap='YlOrRd_r', vmin=0, vmax=0.1)
  
  # Anotasi nilai p-value
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
  ```

### 7. `plot_interaction(data, factor1, factor2, value_col)`

Fungsi untuk memvisualisasikan dan menguji interaksi antar faktor dalam ANOVA faktorial.

- **Contoh kode ANOVA faktorial**:
  ```python
  # Membuat model ANOVA faktorial
  model = ols(f'{value_col} ~ C({factor1}) * C({factor2})', data=data).fit()
  anova_table = sm.stats.anova_lm(model, typ=2)
  
  # Mengecek signifikansi interaksi
  interaction_p = anova_table.loc[f'C({factor1}):C({factor2})', 'PR(>F)']
  
  # Interpretasi hasil
  if interaction_p < 0.05:
      print(f"Interaksi antara {factor1} dan {factor2} signifikan (p = {interaction_p:.4f})")
  else:
      print(f"Interaksi antara {factor1} dan {factor2} tidak signifikan (p = {interaction_p:.4f})")
  ```

### 8. `plot_all_assumptions(data, group_col, value_col)`

Fungsi wrapper yang menjalankan semua fungsi analisis di atas secara berurutan.

- **Contoh kode struktur wrapper**:
  ```python
  def plot_all_assumptions(data, group_col, value_col):
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
  ```

## Perhitungan Statistik

### ANOVA

ANOVA diimplementasikan menggunakan fungsi `ols` dan `anova_lm` dari statsmodels. Formula yang digunakan adalah:

```python
# Membuat model dengan formula
model = ols(f'{value_col} ~ C({group_col})', data=data).fit()

# Menghitung tabel ANOVA
anova_table = sm.stats.anova_lm(model, typ=2)

# Contoh output tabel ANOVA
"""
             df        sum_sq     mean_sq         F    PR(>F)
C(group)    2.0    48.372171  24.186086  6.701906  0.002411
Residual   57.0   205.746283   3.609584       NaN       NaN
"""

# Extraksi hasil untuk interpretasi
f_value = anova_table.loc[f'C({group_col})', 'F']
p_value = anova_table.loc[f'C({group_col})', 'PR(>F)']

# Interpretasi
if p_value < 0.05:
    print(f"Terdapat perbedaan signifikan antar grup (F={f_value:.4f}, p={p_value:.4f})")
else:
    print(f"Tidak terdapat perbedaan signifikan antar grup (F={f_value:.4f}, p={p_value:.4f})")
```

### Tukey HSD (Honestly Significant Difference)

Uji post-hoc Tukey HSD digunakan untuk perbandingan berpasangan setelah ANOVA menunjukkan hasil signifikan:

```python
# Melakukan uji Tukey HSD
tukey = pairwise_tukeyhsd(endog=data[value_col], groups=data[group_col], alpha=0.05)

# Contoh output hasil Tukey HSD
"""
   Multiple Comparison of Means - Tukey HSD, FWER=0.05   
=======================================================
  group1    group2   meandiff  lower   upper  reject
-------------------------------------------------------
Group 1  Group 2    -2.0      -3.5    -0.5   True 
Group 1  Group 3    -4.0      -5.5    -2.5   True 
Group 2  Group 3    -2.0      -3.5    -0.5   True 
-------------------------------------------------------
"""

# Interpretasi per perbandingan
print("Hasil perbandingan berpasangan:")
for i, group1 in enumerate(tukey.groupsunique):
    for j, group2 in enumerate(tukey.groupsunique[i+1:], i+1):
        idx = tukey._find_column_index((group1, group2))
        if idx is not None:
            mean_diff = tukey.meandiffs[idx]
            p_value = tukey.pvalues[idx]
            signif = "signifikan" if p_value < 0.05 else "tidak signifikan"
            print(f"{group1} vs {group2}: perbedaan = {mean_diff:.4f}, p = {p_value:.4f}, {signif}")
```

### Uji Asumsi

1. **Homogenitas Varians**: 
   - Uji Levene menggunakan `stats.levene`
   
   ```python
   # Contoh lengkap
   groups = [data[data[group_col] == group][value_col].values for group in data[group_col].unique()]
   lev_stat, lev_p = stats.levene(*groups)
   print(f"Levene's Test: Statistic = {lev_stat:.4f}, p-value = {lev_p:.4f}")
   print("Kesimpulan: Varians", "homogen" if lev_p > 0.05 else "tidak homogen")
   ```
   
   - Breusch-Pagan menggunakan `het_breuschpagan` dari statsmodels
   
   ```python
   # Contoh lengkap
   model = ols(f'{value_col} ~ C({group_col})', data=data).fit()
   bp_test = het_breuschpagan(model.resid, model.model.exog)
   bp_stat, bp_p, _, _ = bp_test
   print(f"Breusch-Pagan Test: Statistic = {bp_stat:.4f}, p-value = {bp_p:.4f}")
   print("Kesimpulan: Varians", "homoskedastik" if bp_p > 0.05 else "heteroskedastik")
   ```

2. **Normalitas**: 
   - Uji Shapiro-Wilk menggunakan `stats.shapiro`
   
   ```python
   # Contoh untuk residual keseluruhan
   model = ols(f'{value_col} ~ C({group_col})', data=data).fit()
   residuals = model.resid
   sw_stat, sw_p = stats.shapiro(residuals)
   print(f"Shapiro-Wilk Test (residual): Statistic = {sw_stat:.4f}, p-value = {sw_p:.4f}")
   print("Kesimpulan: Residual", "normal" if sw_p > 0.05 else "tidak normal")
   
   # Contoh per grup
   for group in data[group_col].unique():
       group_data = data[data[group_col] == group][value_col]
       sw_stat, sw_p = stats.shapiro(group_data)
       print(f"Grup {group} - Shapiro-Wilk: Statistic = {sw_stat:.4f}, p-value = {sw_p:.4f}")
       print(f"Kesimpulan: Grup {group}", "normal" if sw_p > 0.05 else "tidak normal")
   ```

3. **Independensi**:
   - Durbin-Watson menggunakan `durbin_watson` dari statsmodels
   
   ```python
   # Contoh lengkap
   model = ols(f'{value_col} ~ C({group_col})', data=data).fit()
   residuals = model.resid
   dw_stat = durbin_watson(residuals)
   print(f"Durbin-Watson Statistic: {dw_stat:.4f}")
   
   if dw_stat < 1.5:
       print("Kesimpulan: Kemungkinan ada autokorelasi positif (DW < 1.5)")
   elif dw_stat > 2.5:
       print("Kesimpulan: Kemungkinan ada autokorelasi negatif (DW > 2.5)")
   else:
       print("Kesimpulan: Tidak ada autokorelasi (DW sekitar 2)")
   ```

## Visualisasi Data

Program menghasilkan berbagai visualisasi untuk membantu interpretasi hasil:

1. **Boxplot dan Residual Plot**: Untuk memeriksa homogenitas varians
2. **Histogram dan QQ-Plot**: Untuk memeriksa normalitas
3. **Plot Residual vs Urutan**: Untuk memeriksa independensi
4. **Bar Plot dan Point Plot**: Untuk visualisasi rata-rata dan interval kepercayaan
5. **Mean Difference Plot dan Heatmap**: Untuk visualisasi hasil post-hoc
6. **Interaction Plot**: Untuk visualisasi interaksi antar faktor

Semua visualisasi disimpan ke folder `output/plot/` dalam format PNG.

## Penggunaan

Program ini dirancang untuk:
1. Mengambil data dari file Excel di folder input, khususnya `semua_tabel.xlsx`
2. Memproses data eksperimental koefisien serap bunyi (NRC)
3. Melakukan analisis ANOVA berdasarkan faktor-faktor: Komposisi, Kompaksi, dan Cavity
4. Menghasilkan visualisasi yang diperlukan untuk interpretasi hasil

Jika terjadi kesalahan dalam memuat data, program akan menggunakan data simulasi sebagai contoh penggunaan.

### Contoh Kode Pembacaan Data dari Excel:

```python
# Membaca data dari file Excel
try:
    # Data untuk tabel ANOVA
    anova_data = pd.read_excel('input/semua_tabel.xlsx', sheet_name='tabel_4.6')
    
    # Data eksperimental
    experiment_data = pd.read_excel('input/semua_tabel.xlsx', sheet_name='tabel_4.1')
    
    # Mengkonversi data eksperimen ke format yang sesuai untuk analisis
    df_long = pd.melt(experiment_data, 
                     id_vars=['Komposisi', 'Kompaksi', 'Cavity (mm)'],
                     value_vars=['NRC1', 'NRC2', 'NRC3', 'NRC4', 'NRC5'],
                     var_name='Replikasi', 
                     value_name='NRC')
    
    # Menggunakan data untuk analisis
    plot_all_assumptions(df_long, 'Komposisi', 'NRC')
    
except Exception as e:
    print(f"Error saat memuat atau memproses data: {e}")
    print("Menggunakan data simulasi sebagai contoh")
    contoh_penggunaan() 