import pandas as pd
import os
from openpyxl.styles import NamedStyle

# --- CONFIGURATION ---
# 1. Folder on C: where your NEW CSV is located
input_folder = r'C:\Users\LENOVO\Documents\8020\py_scripts\gg_sold_report\input'

# 2. Folder on G: where your EXISTING Excel files are located
gdrive_folder = r'G:\.shortcut-targets-by-id\1nJVflP2GIzXFjArMBwqzP7HVVHPdhTs1\Client folders\gghomesla\Sold Report\2026' 

# 3. Where you want to save the final EXCEL file
output_path = r'C:\Users\LENOVO\Documents\8020\py_scripts\gg_sold_report\output\Processed_Properties.xlsx'

def run_processor():
    # --- STEP 1: Find the CSV file in the C: drive folder ---
    csv_files = [f for f in os.listdir(input_folder) if f.endswith('.csv')]
    
    if not csv_files:
        print(f"Error: No CSV files found in {input_folder}")
        return
    
    # Process the first CSV found
    csv_path = os.path.join(input_folder, csv_files[0])
    print(f"Reading new data from: {csv_path}")
    
    try:
        df_new = pd.read_csv(csv_path)
    except PermissionError:
        print("Error: The CSV file is open in another program. Please close it.")
        return

    # --- STEP 2: Scan G: Drive for existing property_ids ---
    existing_ids = set()
    print("Scanning G: Drive for existing records...")
    
    if os.path.exists(gdrive_folder):
        excel_files = [f for f in os.listdir(gdrive_folder) if f.endswith(('.xlsx', '.xls'))]
        for file in excel_files:
            try:
                temp_df = pd.read_excel(os.path.join(gdrive_folder, file), usecols=['property_id'])
                existing_ids.update(temp_df['property_id'].dropna().unique())
            except Exception as e:
                print(f"Skipping {file} due to error or missing 'property_id' column.")
    else:
        print("Warning: G: Drive folder not found. Skipping cross-reference.")

    # --- STEP 3: Deduplicate and Filter ---
    initial_count = len(df_new)
    
    # Remove duplicates within the new file itself
    df_new = df_new.drop_duplicates(subset=['property_id'], keep='first')
    
    # Remove IDs that already exist in G: Drive
    unique_df = df_new[~df_new['property_id'].isin(existing_ids)].copy()
    
    # Calculate percentage of removed/repeated values
    removed_count = initial_count - len(unique_df)
    repeat_percent = (removed_count / initial_count) * 100 if initial_count > 0 else 0
    print(f"Percentage of repeated/existing values: {repeat_percent:.2f}%")

    # --- STEP 4: Format Data Types ---
    # Convert money columns
    for col in ['price_sold', 'total_value']:
        if col in unique_df.columns:
            unique_df[col] = pd.to_numeric(unique_df[col], errors='coerce')

    # Convert date column
    if 'sold_date' in unique_df.columns:
        unique_df['sold_date'] = pd.to_datetime(unique_df['sold_date'], errors='coerce')

    # --- STEP 5: Save as Formatted Excel ---
    print(f"Saving formatted Excel to: {output_path}")
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        unique_df.to_excel(writer, index=False, sheet_name='Cleaned_Data')
        
        workbook = writer.book
        worksheet = writer.sheets['Cleaned_Data']
        
        # Define Excel Formats
        money_fmt = '$#,##0.00'
        date_fmt = 'yyyy-mm-dd'
        
        for col_num, column_title in enumerate(unique_df.columns, 1):
            # Apply Money Format
            if column_title in ['price_sold', 'total_value']:
                for row in range(2, len(unique_df) + 2):
                    cell = worksheet.cell(row=row, column=col_num)
                    cell.number_format = money_fmt
            
            # Apply Date Format
            elif column_title == 'sold_date':
                for row in range(2, len(unique_df) + 2):
                    cell = worksheet.cell(row=row, column=col_num)
                    cell.number_format = date_fmt

    print("Process Complete!")

if __name__ == "__main__":
    run_processor()