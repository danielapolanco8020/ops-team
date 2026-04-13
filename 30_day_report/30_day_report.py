import pandas as pd
import os

# The specific columns you want to keep in the final report
columns_to_keep = [
    'FOLIO', 'APN ADDRESS', 'CITY', 'STATE', 'ZIP', 'COUNTY', 'ACTION PLANS', 
    'PROPERTY STATUS', 'SCORE', 'LIKELY DEAL SCORE', 'BUYBOX SCORE', 
    'PROPERTY TYPE', 'VALUE', 'LINK PROPERTIES', 'TAGS', 'HIDDENGEMS', 
    'ABSENTEE', 'HIGH EQUITY', 'DOWNSIZING', 'PRE-FORECLOSURE', 'VACANT', 
    '55+', 'ESTATE', 'INTER FAMILY TRANSFER', 'DIVORCE', 'TAXES', 'PROBATE', 
    'LOW CREDIT', 'CODE VIOLATIONS', 'BANKRUPTCY', 'LIENS CITY/COUNTY', 
    'LIENS OTHER', 'LIENS UTILITY', 'LIENS HOA', 'LIENS MECHANIC', 
    'POOR CONDITION', 'EVICTION', '30-60 DAYS', 'JUDGEMENT', 
    'DEBT COLLECTION', 'DEFAULT RISK'
]

# The sub-set of columns that need to be converted to 1 or 0
binary_cols = [
    'HIDDENGEMS', 'ABSENTEE', 'HIGH EQUITY', 'DOWNSIZING', 'PRE-FORECLOSURE', 
    'VACANT', '55+', 'ESTATE', 'INTER FAMILY TRANSFER', 'DIVORCE', 'TAXES', 
    'PROBATE', 'LOW CREDIT', 'CODE VIOLATIONS', 'BANKRUPTCY', 'LIENS CITY/COUNTY', 
    'LIENS OTHER', 'LIENS UTILITY', 'LIENS HOA', 'LIENS MECHANIC', 'POOR CONDITION', 
    'EVICTION', '30-60 DAYS', 'JUDGEMENT', 'DEBT COLLECTION', 'DEFAULT RISK'
]

def process_atlas_files(directory_path):
    all_summaries = []

    for filename in os.listdir(directory_path):
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            filepath = os.path.join(directory_path, filename)
            df = pd.read_excel(filepath)
            
            # 1. Filter for "30 DAYS" in Action Plans
            df_filtered = df[df['ACTION PLANS'].astype(str).str.upper() == '30 DAYS'].copy()
            
            if df_filtered.empty:
                continue

            # 2. Extract Metadata from filename: "2025-11-10 ATLASRESIDENTIAL 16K Sms"
            parts = filename.replace('.xlsx', '').replace('.xls', '').split(' ')
            file_date = parts[0]  # Gets "2025-11-10"
            channel = parts[-1].upper()  # Gets "SMS" (last word)

            # 3. Clean and Transform columns
            for col in binary_cols:
                if col in df_filtered.columns:
                    # Fill NaNs with 0, then convert values >= 1 to 1, else 0
                    df_filtered[col] = df_filtered[col].fillna(0).apply(lambda x: 1 if x >= 1 else 0)
                else:
                    df_filtered[col] = 0

            # 4. Add the new metadata columns
            df_filtered['DATE'] = file_date
            df_filtered['CHANNEL'] = channel

            # 5. Keep ONLY the requested columns + our new ones
            final_cols = columns_to_keep + ['DATE', 'CHANNEL']
            # Only keep columns that actually exist to avoid errors
            existing_cols = [c for c in final_cols if c in df_filtered.columns]
            df_final = df_filtered[existing_cols]
            
            all_summaries.append(df_final)

    if all_summaries:
        output = pd.concat(all_summaries, ignore_index=True)
        output.to_excel("Final_Property_Summary.xlsx", index=False)
        print(f"Success! Processed {len(all_summaries)} files.")
    else:
        print("No data matched the '30 DAYS' criteria.")

# --- EXECUTION ---
# Replace the path below with your actual folder path
process_atlas_files(r"C:\Users\LENOVO\Documents\8020\py_scripts\30_day_report")