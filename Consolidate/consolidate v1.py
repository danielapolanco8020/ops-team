import os
import pandas as pd
import glob
import re

input_folder = 'Processed_Files'
output_folder = 'Consolidated_Files'
keywords = ["Direct Mail", "Cold Calling", "SMS"]

if not os.path.exists(output_folder):
    os.makedirs(output_folder)

def format_count_to_k(count):
    k_value = count / 1000
    if k_value.is_integer():
        return f"{int(k_value)}K"
    else:
        return f"{round(k_value, 1)}K"

def construct_new_filename(base_filename, new_total_k):
    pattern = r'\s\d+(\.\d+)?K\s'
    replacement = f" {new_total_k} "
    new_name = re.sub(pattern, replacement, base_filename)
    return new_name

print(f"Buscando archivos en: {os.path.abspath(input_folder)}")
print("--- Iniciando Consolidación ---\n")

summary_totals = {}

for keyword in keywords:
    search_pattern = os.path.join(input_folder, f"*{keyword}*.xlsx")
    files = glob.glob(search_pattern)
    
    if not files:
        print(f"⚠️ No se encontraron archivos Excel (.xlsx) para: '{keyword}'")
        continue
        
    print(f"Procesando '{keyword}'... ({len(files)} archivos encontrados)")
    
    dataframes = []
    
    try:
        for file in files:
            df = pd.read_excel(file, dtype=str) 
            dataframes.append(df)
            
        consolidated_df = pd.concat(dataframes, ignore_index=True)

        cols_to_sort = ['BUYBOX SCORE', 'LIKELY DEAL SCORE', 'SCORE']
        
        for col in cols_to_sort:
            if col in consolidated_df.columns:
                consolidated_df[col] = pd.to_numeric(consolidated_df[col], errors='coerce')

        consolidated_df = consolidated_df.sort_values(
            by=cols_to_sort, 
            ascending=[False, False, False]
        )
        
        total_rows = len(consolidated_df)
        summary_totals[keyword] = total_rows
        
        total_k_str = format_count_to_k(total_rows)
        first_filename = os.path.basename(files[0])
        new_filename = construct_new_filename(first_filename, total_k_str)
        
        output_path = os.path.join(output_folder, new_filename)
        consolidated_df.to_excel(output_path, index=False)
        
        print(f" -> ✅ Guardado: {new_filename}")
        
    except Exception as e:
        print(f" -> ❌ Error procesando {keyword}: {e}")

print("\n" + "="*40)
print("RESUMEN FINAL DE PROPIEDADES")
print("="*40)

total_general = 0
for category, count in summary_totals.items():
    print(f"{category:<15}: {count:,.0f} propiedades")
    total_general += count

print("-" * 40)
print(f"TOTAL GENERAL  : {total_general:,.0f} propiedades")
print("="*40)