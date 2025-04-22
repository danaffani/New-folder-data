import pandas as pd
import numpy as np
from scipy import stats

def compare_anova_results():
    try:
        excel_file = "input/semua_tabel.xlsx"
        reference_table = pd.read_excel(excel_file, sheet_name='tabel_4.6')
        
        experiment_data = pd.read_excel('output/nomor3_tambahan.xlsx', sheet_name='RBD')
        
        def calculate_anova_from_data(data):
            n_total = len(data)
            n_rep = 3  
            n_treatments = 8  
            
            grand_mean = data['Nilai Respon'].mean()
            ss_total = sum((data['Nilai Respon'] - grand_mean) ** 2)
            
            treatment_means = data.groupby(['A', 'B', 'C'])['Nilai Respon'].mean()
            ss_treatment = sum((treatment_means - grand_mean) ** 2) * n_rep
            
            def calculate_main_effect_ss(factor):
                factor_means = data.groupby(factor)['Nilai Respon'].mean()
                return sum((factor_means - grand_mean) ** 2) * (n_total / 2)
            
            ss_A = calculate_main_effect_ss('A')  
            ss_B = calculate_main_effect_ss('B')  
            ss_C = calculate_main_effect_ss('C')  
            
            def calculate_interaction_ss(factor1, factor2):
                interaction_means = data.groupby([factor1, factor2])['Nilai Respon'].mean()
                return sum((interaction_means - grand_mean) ** 2) * (n_total / 4) - \
                       calculate_main_effect_ss(factor1) - calculate_main_effect_ss(factor2)
            
            ss_AB = calculate_interaction_ss('A', 'B')  
            ss_AC = calculate_interaction_ss('A', 'C')  
            ss_BC = calculate_interaction_ss('B', 'C')  
            
            ss_ABC = ss_treatment - (ss_A + ss_B + ss_C + ss_AB + ss_AC + ss_BC)
            
            ss_error = ss_total - ss_treatment
            
            df_treatment = n_treatments - 1
            df_main = 1  
            df_interaction = 1  
            df_error = n_total - n_treatments
            df_total = n_total - 1
            
            ms_treatment = ss_treatment / df_treatment
            ms_A = ss_A / df_main
            ms_B = ss_B / df_main
            ms_C = ss_C / df_main
            ms_AB = ss_AB / df_interaction
            ms_AC = ss_AC / df_interaction
            ms_BC = ss_BC / df_interaction
            ms_ABC = ss_ABC / df_interaction
            ms_error = ss_error / df_error
            
            f_A = ms_A / ms_error
            f_B = ms_B / ms_error
            f_C = ms_C / ms_error
            f_AB = ms_AB / ms_error
            f_AC = ms_AC / ms_error
            f_BC = ms_BC / ms_error
            f_ABC = ms_ABC / ms_error
            
            f_crit = stats.f.ppf(0.95, df_main, df_error)
            
            results = pd.DataFrame([
                {'Faktor': 'Perlakuan', 'df': df_treatment, 'SS': ss_treatment, 'MS': ms_treatment},
                {'Faktor': 'Komposisi', 'df': df_main, 'SS': ss_A, 'MS': ms_A, 'Fhit': f_A, 'Ftab': f_crit, 
                 'Keterangan': 'Signifikan' if f_A > f_crit else 'Tidak Signifikan'},
                {'Faktor': 'Kompaksi', 'df': df_main, 'SS': ss_B, 'MS': ms_B, 'Fhit': f_B, 'Ftab': f_crit,
                 'Keterangan': 'Signifikan' if f_B > f_crit else 'Tidak Signifikan'},
                {'Faktor': 'Cavity', 'df': df_main, 'SS': ss_C, 'MS': ms_C, 'Fhit': f_C, 'Ftab': f_crit,
                 'Keterangan': 'Signifikan' if f_C > f_crit else 'Tidak Signifikan'},
                {'Faktor': 'Komposisi*Kompaksi', 'df': df_interaction, 'SS': ss_AB, 'MS': ms_AB, 'Fhit': f_AB, 'Ftab': f_crit,
                 'Keterangan': 'Signifikan' if f_AB > f_crit else 'Tidak Signifikan'},
                {'Faktor': 'Komposisi*Cavity', 'df': df_interaction, 'SS': ss_AC, 'MS': ms_AC, 'Fhit': f_AC, 'Ftab': f_crit,
                 'Keterangan': 'Signifikan' if f_AC > f_crit else 'Tidak Signifikan'},
                {'Faktor': 'Kompaksi*Cavity', 'df': df_interaction, 'SS': ss_BC, 'MS': ms_BC, 'Fhit': f_BC, 'Ftab': f_crit,
                 'Keterangan': 'Signifikan' if f_BC > f_crit else 'Tidak Signifikan'},
                {'Faktor': 'Seluruh Faktor', 'df': df_interaction, 'SS': ss_ABC, 'MS': ms_ABC, 'Fhit': f_ABC, 'Ftab': f_crit,
                 'Keterangan': 'Signifikan' if f_ABC > f_crit else 'Tidak Signifikan'},
                {'Faktor': 'Galat', 'df': df_error, 'SS': ss_error, 'MS': ms_error},
                {'Faktor': 'Total', 'df': df_total, 'SS': ss_total}
            ])
            
            return results
        
        calculated_results = calculate_anova_from_data(experiment_data)
        
        comparison_data = []
        for idx, ref_row in reference_table.iterrows():
            calc_row = calculated_results[calculated_results['Faktor'] == ref_row['Faktor']].iloc[0]
            
            comparison_data.append({
                'Faktor': ref_row['Faktor'],
                'df_ref': ref_row['df'],
                'df_calc': calc_row['df'],
                'SS_ref': ref_row['SS'],
                'SS_calc': calc_row['SS'],
                'MS_ref': ref_row['MS'],
                'MS_calc': calc_row['MS'],
                'Fhit_ref': ref_row['Fhit'] if 'Fhit' in ref_row else None,
                'Fhit_calc': calc_row['Fhit'] if 'Fhit' in calc_row else None,
                'Ftab_ref': ref_row['Ftab'] if 'Ftab' in ref_row else None,
                'Ftab_calc': calc_row['Ftab'] if 'Ftab' in calc_row else None,
                'Keterangan_ref': ref_row['Keterangan'] if 'Keterangan' in ref_row else None,
                'Keterangan_calc': calc_row['Keterangan'] if 'Keterangan' in calc_row else None,
                'Match?': 'Ya' if (
                    'Keterangan' in ref_row and 
                    'Keterangan' in calc_row and 
                    ref_row['Keterangan'] == calc_row['Keterangan']
                ) else 'Tidak'
            })
        
        comparison_df = pd.DataFrame(comparison_data)
        
        with pd.ExcelWriter('output/nomor8_tambahan.xlsx') as writer:
            calculated_results.to_excel(writer, sheet_name='Hasil ANOVA', index=False)
            
            comparison_df.to_excel(writer, sheet_name='Perbandingan', index=False)
            
            reference_table.to_excel(writer, sheet_name='Tabel 4.6 Reference', index=False)
        
        print("\nPerbandingan ANOVA telah disimpan di output/nomor8_tambahan.xlsx")
        print("\nRingkasan Perbandingan:")
        print(comparison_df[['Faktor', 'Match?']])
        
    except Exception as e:
        print(f"Terjadi error: {str(e)}")
        import traceback
        print("\nDetail error:")
        print(traceback.format_exc())

if __name__ == "__main__":
    compare_anova_results() 