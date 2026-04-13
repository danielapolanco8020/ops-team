import pandas as pd
import glob
import os
import csv

def is_in_canada(df):
    # List of Canadian province and territory codes
    canadian_provinces = {
        'AB', 'BC', 'MB', 'NB', 'NL', 'NS', 'NT', 'NU', 'ON', 'PE', 'QC', 'SK', 'YT'
    }
    
    # Clean MailingState: convert to uppercase, strip whitespace, handle NaN
    df['MailingStateClean'] = df['MailingState'].astype(str).str.strip().str.upper()
    
    # Check if MailingState is in the list of Canadian provinces
    df['IsInCanada'] = df['MailingStateClean'].isin(canadian_provinces)
    
    return df

def find_column(df, possible_names):
    # Find a column matching any of the possible names (case-insensitive, ignore spaces/underscores/hyphens)
    for col in df.columns:
        if col.lower().replace(' ', '').replace('_', '').replace('-', '') in \
           [name.lower().replace(' ', '').replace('_', '').replace('-', '') for name in possible_names]:
            return col
    return None

def detect_delimiter(file_path):
    # Use csv.Sniffer to detect the delimiter
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            sample = f.read(1024)  # Read first 1024 bytes
            sniffer = csv.Sniffer()
            return sniffer.sniff(sample).delimiter
    except Exception as e:
        print(f"Could not detect delimiter for {file_path}: {str(e)}. Defaulting to comma.")
        return ','

def inspect_file(file_path, num_lines=5):
    # Print the first few lines of the file for debugging
    print(f"\nInspecting first {num_lines} lines of {file_path}:")
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            for i, line in enumerate(f, 1):
                if i > num_lines:
                    break
                print(f"Line {i}: {line.strip()}")
    except Exception as e:
        print(f"Error reading {file_path}: {str(e)}")

def process_csv_files(input_files, output_file):
    # Initialize an empty list to store DataFrames with Canadian locations
    canadian_dfs = []
    
    # Process each CSV file
    for file in input_files:
        try:
            # Inspect the file content
            inspect_file(file)
            
            # Detect the delimiter
            delimiter = detect_delimiter(file)
            print(f"Detected delimiter for {file}: '{delimiter}'")
            
            # Try reading the CSV with the detected delimiter
            try:
                df = pd.read_csv(file, encoding='utf-8', sep=delimiter, on_bad_lines='warn', quoting=csv.QUOTE_ALL)
            except Exception as e:
                print(f"Failed to read {file} with delimiter '{delimiter}': {str(e)}")
                # Try without header and infer columns
                df = pd.read_csv(file, encoding='utf-8', sep=delimiter, header=None, on_bad_lines='warn', quoting=csv.QUOTE_ALL)
                # Attempt to find MailingCity and MailingState in the first row
                first_row = df.iloc[0].astype(str).str.lower().replace(' ', '').replace('_', '').replace('-', '')
                city_col_idx = None
                state_col_idx = None
                for idx, val in first_row.items():
                    if val in [name.lower().replace(' ', '').replace('_', '').replace('-', '') 
                              for name in ['MailingCity', 'City', 'mailing_city', 'Mailing City', 'mailing-city']]:
                        city_col_idx = idx
                    if val in [name.lower().replace(' ', '').replace('_', '').replace('-', '') 
                              for name in ['MailingState', 'State', 'Province', 'Mailing_State', 'mailing_state', 'Mailing State', 'mailing-state']]:
                        state_col_idx = idx
                if city_col_idx is not None and state_col_idx is not None:
                    df.columns = [f'col_{i}' for i in range(len(df.columns))]
                    df.columns[city_col_idx] = 'MailingCity'
                    df.columns[state_col_idx] = 'MailingState'
                    df = df.drop(0).reset_index(drop=True)
                else:
                    print(f"Warning: {file} has no recognizable 'MailingCity' or 'MailingState' in first row. Skipping.")
                    continue
            
            # Debug: Print column names
            print(f"Columns in {file}: {list(df.columns)}")
            
            # Find MailingCity and MailingState columns
            city_col = find_column(df, ['MailingCity', 'City', 'mailing_city', 'Mailing City', 'mailing-city'])
            state_col = find_column(df, ['MailingState', 'State', 'Province', 'Mailing_State', 'mailing_state', 'Mailing State', 'mailing-state'])
            
            # Verify required columns exist
            if not city_col or not state_col:
                print(f"Warning: {file} is missing 'MailingCity' ({city_col}) or 'MailingState' ({state_col}) columns. Skipping.")
                continue
            
            # Rename columns to standard names for consistency
            df = df.rename(columns={city_col: 'MailingCity', state_col: 'MailingState'})
            
            # Debug: Print unique MailingState values
            unique_states = df['MailingState'].astype(str).unique()
            print(f"Unique MailingState values in {file}: {unique_states}")
            
            # Check for Canadian locations
            df = is_in_canada(df)
            
            # Filter rows where IsInCanada is True
            canadian_rows = df[df['IsInCanada']].copy()
            
            # Drop temporary columns to maintain original columns
            canadian_rows = canadian_rows.drop(columns=['IsInCanada', 'MailingStateClean'])
            
            # Add a column to track the source file
            canadian_rows['SourceFile'] = os.path.basename(file)
            
            # Append to the list if there are valid rows
            if not canadian_rows.empty:
                canadian_dfs.append(canadian_rows)
                print(f"Processed {file}: {len(canadian_rows)} Canadian rows found.")
            else:
                print(f"Processed {file}: No Canadian rows found.")
                
        except Exception as e:
            print(f"Error processing {file}: {str(e)}")
    
    # Merge all valid rows into a single DataFrame
    if canadian_dfs:
        merged_df = pd.concat(canadian_dfs, ignore_index=True)
        
        # Save the merged DataFrame to a CSV file
        merged_df.to_csv(output_file, index=False)
        print(f"Merged {len(merged_df)} Canadian rows into {output_file}.")
    else:
        print("No Canadian rows found in any files. No output file created.")
        merged_df = pd.DataFrame()
    
    return merged_df

# Example usage
if __name__ == "__main__":
    # Specify the input CSV files path
    input_path = r"C:\Users\LENOVO\Documents\8020\py_scripts\Amerian_HBs_canadian\input\*.csv"
    input_files = glob.glob(input_path)
    
    # Specify the output file
    output_file = r"C:\Users\LENOVO\Documents\8020\py_scripts\Amerian_HBs_canadian\canadian_locations_merged.csv"
    
    # Check if input files exist
    if not input_files:
        print(f"No CSV files found in {input_path}")
    else:
        # Process the files and merge Canadian rows
        merged_df = process_csv_files(input_files, output_file)
        
        # Display the merged DataFrame
        if not merged_df.empty:
            print("\nPreview of merged DataFrame:")
            print(merged_df.head())