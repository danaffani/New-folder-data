import pandas as pd
import numpy as np
from scipy import stats

def compare_anova_results():
    try:
        reference_file = "input/semua_tabel.xlsx"
        reference_table = pd.read_excel(reference_file, sheet_name='tabel_4.6')
        
        data_file = "input/tabel_koef_Serap_bunyi.xlsx"
        
        experiment_data = pd.read_excel(data_file, sheet_name='CRD')
        
        experiment_data['A'] = experiment_data['Komposisi'].map({"50 : 50": -1, "70 : 30": 0, "90 : 10": 1})
        experiment_data['B'] = experiment_data['Kompaksi'].map({"3 : 4": -1, "4 : 4": 0, "5 : 4": 1})
        experiment_data['C'] = experiment_data['Cavity'].map({"15 mm": -1, "20 mm": 0, "25 mm": 1})
        
        experiment_data.rename(columns={'Rata-rata Koefisien': 'Nilai Respon'}, inplace=True)
        
        all_results = {}
        
        for frekuensi in experiment_data['Frekuensi'].unique():
            data_for_freq = experiment_data[experiment_data['Frekuensi'] == frekuensi].copy()
            
            def calculate_anova_for_frequency(data):
                n_total = len(data)
                n_rep = 1
                n_treatments = 27
                
                grand_mean = data['Nilai Respon'].mean()
                ss_total = sum((data['Nilai Respon'] - grand_mean) ** 2)
                
                treatment_means = data.groupby(['A', 'B', 'C'])['Nilai Respon'].mean()
                ss_treatment = sum((treatment_means - grand_mean) ** 2) * n_rep
                
                def calculate_main_effect_ss(factor):
                    factor_means = data.groupby(factor)['Nilai Respon'].mean()
                    return sum((factor_means - grand_mean) ** 2) * (n_total / len(factor_means))
                
                ss_A = calculate_main_effect_ss('A')  
                ss_B = calculate_main_effect_ss('B')  
                ss_C = calculate_main_effect_ss('C')  
                
                def calculate_interaction_ss(factor1, factor2):
                    interaction_means = data.groupby([factor1, factor2])['Nilai Respon'].mean()
                    n_combinations = len(interaction_means)
                    return sum((interaction_means - grand_mean) ** 2) * (n_total / n_combinations) - \
                           calculate_main_effect_ss(factor1) - calculate_main_effect_ss(factor2)
                
                ss_AB = calculate_interaction_ss('A', 'B')  
                ss_AC = calculate_interaction_ss('A', 'C')  
                ss_BC = calculate_interaction_ss('B', 'C')  
                
                ss_ABC = ss_treatment - (ss_A + ss_B + ss_C + ss_AB + ss_AC + ss_BC)
                
                ss_error = ss_total - ss_treatment
                
                df_A = 2  #3 level faktor A (-1, 0, 1)
                df_B = 2  #3 level faktor B (-1, 0, 1)
                df_C = 2  #3 level faktor C (-1, 0, 1) 
                df_AB = df_A * df_B
                df_AC = df_A * df_C
                df_BC = df_B * df_C
                df_ABC = df_A * df_B * df_C
                df_treatment = n_treatments - 1
                df_error = n_total - n_treatments
                df_total = n_total - 1
                
                ms_A = ss_A / df_A
                ms_B = ss_B / df_B
                ms_C = ss_C / df_C
                ms_AB = ss_AB / df_AB  
                ms_AC = ss_AC / df_AC
                ms_BC = ss_BC / df_BC
                ms_ABC = ss_ABC / df_ABC
                ms_treatment = ss_treatment / df_treatment
                ms_error = ss_error / df_error if df_error > 0 else 0
                
                f_A = ms_A / ms_error if ms_error > 0 else 0
                f_B = ms_B / ms_error if ms_error > 0 else 0
                f_C = ms_C / ms_error if ms_error > 0 else 0
                f_AB = ms_AB / ms_error if ms_error > 0 else 0
                f_AC = ms_AC / ms_error if ms_error > 0 else 0
                f_BC = ms_BC / ms_error if ms_error > 0 else 0
                f_ABC = ms_ABC / ms_error if ms_error > 0 else 0
                
                f_crit_A = stats.f.ppf(0.95, df_A, df_error) if df_error > 0 else 0
                f_crit_B = stats.f.ppf(0.95, df_B, df_error) if df_error > 0 else 0
                f_crit_C = stats.f.ppf(0.95, df_C, df_error) if df_error > 0 else 0
                f_crit_AB = stats.f.ppf(0.95, df_AB, df_error) if df_error > 0 else 0
                f_crit_AC = stats.f.ppf(0.95, df_AC, df_error) if df_error > 0 else 0
                f_crit_BC = stats.f.ppf(0.95, df_BC, df_error) if df_error > 0 else 0
                f_crit_ABC = stats.f.ppf(0.95, df_ABC, df_error) if df_error > 0 else 0
                
                results = pd.DataFrame([
                    {'Faktor': 'Perlakuan', 'df': df_treatment, 'SS': ss_treatment, 'MS': ms_treatment},
                    {'Faktor': 'Komposisi', 'df': df_A, 'SS': ss_A, 'MS': ms_A, 'Fhit': f_A, 'Ftab': f_crit_A, 
                     'Keterangan': 'Signifikan' if f_A > f_crit_A else 'Tidak Signifikan'},
                    {'Faktor': 'Kompaksi', 'df': df_B, 'SS': ss_B, 'MS': ms_B, 'Fhit': f_B, 'Ftab': f_crit_B,
                     'Keterangan': 'Signifikan' if f_B > f_crit_B else 'Tidak Signifikan'},
                    {'Faktor': 'Cavity', 'df': df_C, 'SS': ss_C, 'MS': ms_C, 'Fhit': f_C, 'Ftab': f_crit_C,
                     'Keterangan': 'Signifikan' if f_C > f_crit_C else 'Tidak Signifikan'},
                    {'Faktor': 'Komposisi*Kompaksi', 'df': df_AB, 'SS': ss_AB, 'MS': ms_AB, 'Fhit': f_AB, 'Ftab': f_crit_AB,
                     'Keterangan': 'Signifikan' if f_AB > f_crit_AB else 'Tidak Signifikan'},
                    {'Faktor': 'Komposisi*Cavity', 'df': df_AC, 'SS': ss_AC, 'MS': ms_AC, 'Fhit': f_AC, 'Ftab': f_crit_AC,
                     'Keterangan': 'Signifikan' if f_AC > f_crit_AC else 'Tidak Signifikan'},
                    {'Faktor': 'Kompaksi*Cavity', 'df': df_BC, 'SS': ss_BC, 'MS': ms_BC, 'Fhit': f_BC, 'Ftab': f_crit_BC,
                     'Keterangan': 'Signifikan' if f_BC > f_crit_BC else 'Tidak Signifikan'},
                    {'Faktor': 'Seluruh Faktor', 'df': df_ABC, 'SS': ss_ABC, 'MS': ms_ABC, 'Fhit': f_ABC, 'Ftab': f_crit_ABC,
                     'Keterangan': 'Signifikan' if f_ABC > f_crit_ABC else 'Tidak Signifikan'},
                    {'Faktor': 'Galat', 'df': df_error, 'SS': ss_error, 'MS': ms_error},
                    {'Faktor': 'Total', 'df': df_total, 'SS': ss_total}
                ])
                
                return results
            
            all_results[frekuensi] = calculate_anova_for_frequency(data_for_freq)
        
        comparison_data = []
        if "Rata-rata" in all_results:
            calculated_results = all_results["Rata-rata"]
            
            for idx, ref_row in reference_table.iterrows():
                calc_rows = calculated_results[calculated_results['Faktor'] == ref_row['Faktor']]
                
                if not calc_rows.empty:
                    calc_row = calc_rows.iloc[0]
                    
                    comparison_data.append({
                        'Faktor': ref_row['Faktor'],
                        'df_ref': ref_row.get('df', None),
                        'df_calc': calc_row.get('df', None),
                        'SS_ref': ref_row.get('SS', None),
                        'SS_calc': calc_row.get('SS', None),
                        'MS_ref': ref_row.get('MS', None),
                        'MS_calc': calc_row.get('MS', None),
                        'Fhit_ref': ref_row.get('Fhit', None),
                        'Fhit_calc': calc_row.get('Fhit', None),
                        'Ftab_ref': ref_row.get('Ftab', None),
                        'Ftab_calc': calc_row.get('Ftab', None),
                        'Keterangan_ref': ref_row.get('Keterangan', None),
                        'Keterangan_calc': calc_row.get('Keterangan', None),
                        'Match?': 'Ya' if (
                            'Keterangan' in ref_row and 
                            'Keterangan' in calc_row and 
                            ref_row['Keterangan'] == calc_row['Keterangan']
                        ) else 'Tidak'
                    })
        
        comparison_df = pd.DataFrame(comparison_data)
        
        with pd.ExcelWriter('output/nomor8_tambahan.xlsx') as writer:
            for freq, results in all_results.items():
                sheet_name = freq.replace(" ", "_") if isinstance(freq, str) else f"Frequency_{freq}"
                sheet_name = 'ANOVA_' + sheet_name
                results.to_excel(writer, sheet_name=sheet_name, index=False)
            
            comparison_df.to_excel(writer, sheet_name='Perbandingan', index=False)
            reference_table.to_excel(writer, sheet_name='Tabel 4.6 Reference', index=False)
            
            experiment_data.to_excel(writer, sheet_name='Data_Lengkap', index=False)
        
        print("\nPerbandingan ANOVA telah disimpan di output/nomor8_tambahan.xlsx")
        print(f"Hasil ANOVA untuk {len(all_results)} frekuensi berbeda telah disimpan")
        print(f"Total data yang diolah: {len(experiment_data)} baris")
        print("\nRingkasan Perbandingan dengan referensi:")
        if not comparison_df.empty:
            print(comparison_df[['Faktor', 'Match?']])
        else:
            print("Tidak dapat membuat perbandingan karena data Rata-rata tidak ditemukan")
        
    except Exception as e:
        print(f"Terjadi error: {str(e)}")
        import traceback
        print("\nDetail error:")
        print(traceback.format_exc())

if __name__ == "__main__":
    compare_anova_results() 