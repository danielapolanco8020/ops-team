import pandas as pd
import os
from tqdm import tqdm
import time
import re
import multiprocessing as mp
from functools import partial

def filter_unwanted_names(df, column_name, unwanted_names):
    pattern = re.compile('|'.join(map(re.escape, unwanted_names)), re.IGNORECASE)
    mask = df[column_name].str.contains(pattern, na=False)
    rejected = df[mask].copy()
    rejected['Rejection_Stage'] = 'Unwanted Names'
    rejected['Rejection_Value'] = rejected[column_name]
    return df[~mask], rejected

def filter_owner_types(df, column_name, unwanted_types):
    if column_name not in df.columns:
        print(f"Warning: Column '{column_name}' not found in DataFrame. Skipping owner type filtering.")
        return df, pd.DataFrame(columns=df.columns.tolist() + ['Rejection_Stage', 'Rejection_Value'])
    pattern = re.compile('|'.join(map(re.escape, unwanted_types)), re.IGNORECASE)
    mask = df[column_name].str.contains(pattern, na=False)
    rejected = df[mask].copy()
    rejected['Rejection_Stage'] = 'Unwanted Owner Type'
    rejected['Rejection_Value'] = rejected[column_name]
    return df[~mask], rejected

def phone_count_check(df, file_path, phone_type_cols=None, phone_number_cols=None):
    try:
        file_name = os.path.basename(file_path)
        keywords = ['sms', 'cc commodore', 'cold calling']
        if any(re.search(r'\b{}\b'.format(re.escape(keyword)), file_name, re.IGNORECASE) for keyword in keywords):
            # Cache column identification
            if phone_type_cols is None:
                phone_type_cols = [col for col in df.columns if 'PHONE TYPE' in col.upper()]
            if phone_number_cols is None:
                phone_number_cols = [col for col in df.columns if 'PHONE NUMBER' in col.upper()]
            all_phone_cols = phone_type_cols + phone_number_cols
            
            # Single filter for properties with data
            properties_with_data_mask = df[all_phone_cols].notna().any(axis=1)
            properties_with_data_count = properties_with_data_mask.sum()
            empty_rows_count = len(df) - properties_with_data_count
            
            # Buffer output
            output = [f"\n--- PHONE NUMBER COUNT (CLEANED DATA) ---",
                      f"File Analyzed: {file_path}",
                      f"Properties with Active Phone Numbers: {properties_with_data_count}",
                      f"Properties without Active Phone Numbers: {empty_rows_count}"]
            
            # Skiptrace check with precompiled regex
            if 'TAGS' in df.columns:
                skip_pattern = re.compile(r'\bskip\b', re.IGNORECASE)
                skiptrace_count = len(df[~properties_with_data_mask][df['TAGS'].str.contains(skip_pattern, na=False)])
                output.append(f"Properties without 'SKIPTRACE': {skiptrace_count}")
            else:
                output.append("TAGS column not found, skipping SKIPTRACE check.")
            
            print('\n'.join(output))
        else:
            print(f"File name '{file_name}' does not contain keywords related to SMS or Cold Call.")
    except Exception as e:
        print(f"Error in phone_count_check for {file_path}: {str(e)}")

def save_rejection_summary(rejected_properties, folder_path, output_folder):
    summary_path = os.path.join(output_folder, "Rejection_Summary.xlsx")
    rejected_properties_path = os.path.join(output_folder, "Rejected_Properties.xlsx")
    try:
        if not rejected_properties.empty:
            # Append rejected properties to Rejected_Properties.xlsx
            if os.path.exists(rejected_properties_path):
                existing_rejected = pd.read_excel(rejected_properties_path, engine='openpyxl')
                combined_rejected = pd.concat([existing_rejected, rejected_properties], ignore_index=True)
            else:
                combined_rejected = rejected_properties
            combined_rejected.to_excel(rejected_properties_path, index=False, engine='openpyxl')
            print(f"Rejected properties appended to {rejected_properties_path}")

            # Process rejection summary
            current_files = set(f for f in os.listdir(folder_path) if f.endswith('.xlsx'))
            rejected_properties_filtered = rejected_properties[rejected_properties['Source_File'].isin(current_files)]
            if rejected_properties_filtered.empty:
                print("No rejected properties from current files to summarize.")
                return
            new_summary = rejected_properties_filtered.groupby(['Source_File', 'Rejection_Stage']).size().reset_index(name='Rejected_Count')
            original_counts = {}
            for file in rejected_properties_filtered['Source_File'].unique():
                file_path = os.path.join(folder_path, file)
                try:
                    df = pd.read_excel(file_path, engine='openpyxl')
                    original_counts[file] = len(df)
                except FileNotFoundError:
                    print(f"Warning: File {file} not found in {folder_path}, skipping in summary.")
                    continue
            new_summary['Original_Count'] = new_summary['Source_File'].map(original_counts)
            new_summary['Original_Count'] = new_summary['Original_Count'].fillna(0).astype(int)
            new_summary['Rejection_Percentage'] = (
                new_summary['Rejected_Count'] / new_summary['Original_Count'] * 100
            ).round(2).fillna(0.0)
            if os.path.exists(summary_path):
                existing_summary = pd.read_excel(summary_path, engine='openpyxl')
                print(f"Loaded {len(existing_summary)} existing summary rows.")
            else:
                existing_summary = pd.DataFrame(columns=['Source_File', 'Rejection_Stage', 'Rejected_Count', 'Original_Count', 'Rejection_Percentage'])
                print("No existing rejection summary found, starting fresh.")
            existing_summary = existing_summary[~existing_summary['Source_File'].isin(current_files)]
            combined_summary = pd.concat([existing_summary, new_summary], ignore_index=True)
            combined_summary.to_excel(summary_path, index=False, engine='openpyxl')
            print(f"Rejection summary saved at {summary_path}")
    except Exception as e:
        print(f"Error in save_rejection_summary: {str(e)}")

def process_single_file(file, folder_path, output_folder):
    start_time = time.time()
    file_path = os.path.join(folder_path, file)
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
    except FileNotFoundError:
        print(f"Error: File {file} not found, skipping.")
        return None, None, 0, time.time() - start_time
    except Exception as e:
        print(f"Error processing {file}: {str(e)}")
        return None, None, 0, time.time() - start_time
    original_row_count = len(df)
    rejected_rows_list = []
    df['OWNER LAST NAME'] = df['OWNER LAST NAME'].fillna(df['OWNER FULL NAME'])
    df.sort_values(by=['BUYBOX SCORE', 'LIKELY DEAL SCORE', 'SCORE'], ascending=False, inplace=True)
    mask = df['ACTION PLANS'].notna()
    rejected = df[~mask].copy()
    rejected['Rejection_Stage'] = 'Empty Action Plans'
    rejected['Rejection_Value'] = rejected['ACTION PLANS'].astype(str)
    rejected['Source_File'] = file
    rejected_rows_list.append(rejected)
    df = df[mask]
    df['MAILING ADDRESS'] = df['MAILING ADDRESS'].fillna(df['ADDRESS'])
    df['MAILING ZIP'] = df['MAILING ZIP'].fillna(df['ZIP'])
    for subset, stage in [(['MAILING ADDRESS', 'MAILING ZIP'], 'Duplicate Address'),
                          (['OWNER FULL NAME', 'ADDRESS', 'ZIP'], 'Duplicate Owner')]:
        duplicated_mask = df.duplicated(subset=subset, keep=False)
        rejected = df[duplicated_mask].copy()
        rejected['Rejection_Stage'] = stage
        rejected['Rejection_Value'] = rejected.apply(
            lambda row: ' | '.join(str(row[col]) for col in subset), axis=1
        )
        rejected['Source_File'] = file
        rejected_rows_list.append(rejected)
        df = df[~duplicated_mask]
    unwanted_names = [
        "Given Not", "Record", "Available", "Bank ", "Church ", "School", "Cemetery",
        "Not given", "University", "College", "Owner", "Hospital", "County", "City of", "Not Provided Name"
    ]
    df, rejected = filter_unwanted_names(df, 'OWNER FULL NAME', unwanted_names)
    rejected['Source_File'] = file
    rejected_rows_list.append(rejected)
    unwanted_types = ["Non Sellers", "Religious Organization"]
    df, rejected = filter_owner_types(df, 'OWNER TYPE', unwanted_types)
    rejected['Source_File'] = file
    rejected_rows_list.append(rejected)
    df = df.drop(columns=['OWNER TYPE'], errors='ignore')
    print(f"Number of properties in {file}: {len(df)}")
    # Precompute phone-related columns for efficiency
    phone_type_cols = [col for col in df.columns if 'PHONE TYPE' in col.upper()]
    phone_number_cols = [col for col in df.columns if 'PHONE NUMBER' in col.upper()]
    # Run phone_count_check on cleaned DataFrame with precomputed columns
    phone_count_check(df, file_path, phone_type_cols, phone_number_cols)
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    try:
        df.to_excel(os.path.join(output_folder, f"output - {file}"), index=False, engine='openpyxl')
    except Exception as e:
        print(f"Error saving output for {file}: {str(e)}")
    rejected_rows = pd.concat(rejected_rows_list, ignore_index=True) if rejected_rows_list else pd.DataFrame(columns=df.columns.tolist() + ['Rejection_Stage', 'Rejection_Value', 'Source_File'])
    # Save rejected properties and update rejection summary
    save_rejection_summary(rejected_rows, folder_path, output_folder)
    return df, rejected_rows, original_row_count, time.time() - start_time

def process_excel_files_v2(folder_path, output_folder):
    log_output_path = os.path.join(output_folder, "Processing_Log.txt")
    try:
        files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
        if not files:
            print(f"No Excel files found in {folder_path}. Exiting.")
            return
        log_entries = []
        with mp.Pool(processes=mp.cpu_count()) as pool:
            results = list(tqdm(pool.imap_unordered(
                partial(process_single_file, folder_path=folder_path, output_folder=output_folder),
                files), total=len(files), desc="Processing files"))
        for result in results:
            if result is None:
                log_entries.append("0 | 0.00 (File not found)")
                continue
            df, rejected_rows, row_count, proc_time = result
            log_entries.append(f"{row_count} | {proc_time:.2f}")
        with open(log_output_path, 'a') as log_file:
            log_file.write(f"\nProcessing started at {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            log_file.write("Row Count | Processing Time (seconds)\n")
            log_file.write("-" * 40 + "\n")
            for entry in log_entries:
                log_file.write(f"{entry}\n")
    except Exception as e:
        print(f"Error in process_excel_files_v2: {str(e)}")

if __name__ == "__main__":
    try:
        process_excel_files_v2(
            r"C:\Users\LENOVO\Documents\8020\AutomatedManualFulfillment\Income File",
            r"C:\Users\LENOVO\Documents\8020\AutomatedManualFulfillment\Output file"
        )
    except Exception as e:
        print(f"Main execution error: {str(e)}")