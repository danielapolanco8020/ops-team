import pandas as pd
import os
import re

distress_cols = [
    'divorcedistress', 'interfamilydistress', 'interfamilydistressdate', 'downsizingdistress',
    'estatedistress', 'lowincomedistress', 'poorconditiondistress', 'seniordistress',
    'taxdelinquentdistress', 'preforeclosuredistress', 'failedlistingdistress', 'hoadistress',
    'debtcollectiondistress', 'judgmentdistress', 'liencitycountydistress', 'probatedistress',
    'lienutilitydistress', 'lienmechanicaldistress', 'affidavitdistress', 'lienotherdistress',
    'bankruptcydistress', 'thirtysixtydaysdistress', 'evictiondistress', 'violationdistress'
]

def clean_string(val):
    """Deep cleans strings: Uppercase, removes punctuation/extra spaces."""
    if pd.isna(val): return ""
    # Remove everything that isn't a letter or number, then uppercase
    s = str(val).strip().upper()
    return re.sub(r'[^A-Z0-9]', '', s)

def clean_zip(val):
    """Ensures ZIP is a 5-character string (handles leading zeros)."""
    if pd.isna(val): return ""
    s = str(val).split('.')[0].strip() # Remove .0 if it's a float
    return s.zfill(5)[-5:] # Pad with zeros to 5 digits

def get_existing_keys(directory_path):
    existing_keys = set()
    for filename in os.listdir(directory_path):
        if filename.endswith((".xlsx", ".xls")):
            filepath = os.path.join(directory_path, filename)
            try:
                # Load headers to find the address column
                df_headers = pd.read_excel(filepath, nrows=0)
                cols = df_headers.columns.tolist()
                addr_col = 'APN ADDRESS' if 'APN ADDRESS' in cols else 'ADDRESS' if 'ADDRESS' in cols else None
                
                if addr_col and all(c in cols for c in ['CITY', 'STATE', 'ZIP']):
                    # Load and clean data
                    temp_df = pd.read_excel(filepath, usecols=[addr_col, 'CITY', 'STATE', 'ZIP'])
                    
                    keys = (temp_df[addr_col].apply(clean_string) + 
                            temp_df['CITY'].apply(clean_string) + 
                            temp_df['STATE'].apply(clean_string) + 
                            temp_df['ZIP'].apply(clean_zip))
                    existing_keys.update(keys.tolist())
            except Exception as e:
                print(f"Skipping {filename}: {e}")
    return existing_keys

def process_matching_leads(csv_path, excel_folder_path):
    # 1. Load CSV
    df = pd.read_csv(csv_path)

    # 2. Convert Distress columns Y/N -> 1/0
    for col in distress_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.upper().map({'Y': 1, 'N': 0}).fillna(0).astype(int)

    # 3. Create deep-cleaned unique keys
    df['unique_key'] = (df['situsfullstreetaddress'].apply(clean_string) + 
                        df['situscity'].apply(clean_string) + 
                        df['situsstate'].apply(clean_string) + 
                        df['situszip5'].apply(clean_zip))

    # 4. Scan history
    print("Building deep-cleaned historical database...")
    existing_keys = get_existing_keys(excel_folder_path)

    # 5. Filter: Keep only matches
    original_count = len(df)
    df_matches = df[df['unique_key'].isin(existing_keys)].copy()
    
    df_matches = df_matches.drop(columns=['unique_key'])

    # 6. Save
    output_name = "Matched_Historical_Leads_Enhanced.csv"
    df_matches.to_csv(output_name, index=False)
    
    print("-" * 30)
    print(f"Process Complete!")
    print(f"Total leads in CSV: {original_count}")
    print(f"Matches found:      {len(df_matches)}")
    print(f"File saved as: {output_name}")

# Paths (Keep as you had them)
new_csv_file = r"C:\Users\LENOVO\Documents\8020\py_scripts\30_day_report\Distress_report\3bff267e-e0b6-4e43-b739-7d4c8f2fdff2 (1).csv"
history_folder = r"C:\Users\LENOVO\Documents\8020\py_scripts\30_day_report\Distress_report\fulfillments"

process_matching_leads(new_csv_file, history_folder)