import os
import pandas as pd
from tqdm import tqdm

# ZIP filter configuration
ZIP_FILTER_FILE_PATH = r"C:\Users\LENOVO\Documents\8020\py_scripts\Benchmark_report\ZIP_benchmark.xlsx"
ZIP_COLUMN_NAME = 'ZIP codes'

def rename_columns(df):
    """
    Rename columns to standardized names after filtering is complete.
    """
    # Define the column mapping (old_name -> new_name)
    column_mapping = {
        'zip': 'ZIP',
        'SaleDate': 'LAST SALE DATE',
        'FIPS': 'COUNTY',  # Changed from 'County' to 'FIPS' since that's the actual column name
        'LotSize': 'LOT SIZE',
        'YearBuilt': 'YEAR BUILT',
        'LivingArea': 'LIVING SQFT',
        'OwnerType': 'OWNER TYPE',
        'TotalValue': 'VALUE',
        'propertytype': 'PROPERTY TYPE',
        'County': 'County ETL'
    }
    
    # Track which columns were renamed
    renamed_columns = []
    available_columns = df.columns.tolist()
    
    # Apply renaming only for columns that exist
    columns_to_rename = {}
    for old_name, new_name in column_mapping.items():
        if old_name in available_columns:
            columns_to_rename[old_name] = new_name
            renamed_columns.append(f"'{old_name}' -> '{new_name}'")
    
    if columns_to_rename:
        df_renamed = df.rename(columns=columns_to_rename)
        print(f"6. Renamed columns: {', '.join(renamed_columns)}")
        return df_renamed
    else:
        print("6. No columns found to rename")
        print(f"   Available columns: {available_columns}")
        return df

def load_valid_zip_codes(zip_file_path):
    """
    Loads valid ZIP codes from an external Excel file.
    Uses the ZIP_COLUMN_NAME constant to find the correct column.
    Returns a set of valid ZIP codes for faster lookup.
    """
    try:
        zip_df = pd.read_excel(zip_file_path)
        
        # Check if the specified column exists
        if ZIP_COLUMN_NAME not in zip_df.columns:
            print(f"Warning: Column '{ZIP_COLUMN_NAME}' not found in {zip_file_path}")
            print(f"Available columns: {list(zip_df.columns)}")
            print("Proceeding without ZIP code filtering...")
            return None
        
        # Get the specified column and clean the data
        zip_column = zip_df[ZIP_COLUMN_NAME]
        
        # Remove NaN values and convert to string, then normalize
        valid_zips = set()
        for zip_code in zip_column.dropna():
            # Convert to string, strip whitespace, and ensure 5-digit format
            zip_str = str(zip_code).strip()
            # Handle both numeric and string ZIP codes
            if zip_str.isdigit() and len(zip_str) <= 5:
                zip_str = zip_str.zfill(5)  # Pad with leading zeros if needed
            valid_zips.add(zip_str)
        
        print(f"Loaded {len(valid_zips)} valid ZIP codes from ZIP filter file")
        return valid_zips
    except FileNotFoundError:
        print(f"Warning: ZIP filter file not found at {zip_file_path}")
        print("Proceeding without ZIP code filtering...")
        return None
    except Exception as e:
        print(f"Error loading ZIP codes from {zip_file_path}: {e}")
        print("Proceeding without ZIP code filtering...")
        return None

def filter_by_zip_codes(df, valid_zip_codes, zip_column_name='ZIP'):
    """
    Simplified and more reliable ZIP filtering function.
    """
    if valid_zip_codes is None:
        print("No valid ZIP codes provided - skipping ZIP filter")
        return df, 0
    
    # Find ZIP column
    possible_zip_columns = [
        zip_column_name, 'ZIP', 'Zip', 'zip', 'ZIP_CODE', 'ZIP CODE', 
        'Zip Code', 'zip_code', 'ZIPCODE', 'ZipCode', 'zipcode', 'ZIP codes'
    ]
    
    zip_col_found = None
    for col_name in possible_zip_columns:
        if col_name in df.columns:
            zip_col_found = col_name
            break
    
    if zip_col_found is None:
        print(f"No ZIP column found in: {list(df.columns)}")
        return df, 0
    
    print(f"Using ZIP column: '{zip_col_found}'")
    original_count = len(df)
    
    # Simple approach: normalize both sets to 5-digit strings
    def normalize_zip(zip_val):
        if pd.isna(zip_val):
            return ''
        zip_str = str(zip_val).strip()
        # Remove decimal point if present (Excel sometimes converts to float)
        if '.' in zip_str:
            zip_str = zip_str.split('.')[0]
        # Take only first 5 digits if longer
        if zip_str.isdigit():
            zip_str = zip_str[:5].zfill(5)
        return zip_str
    
    # Normalize ZIP codes in the DataFrame
    df_normalized = df.copy()
    df_normalized['_temp_zip'] = df[zip_col_found].apply(normalize_zip)
    
    # Filter using the normalized ZIP codes
    matches = df_normalized['_temp_zip'].isin(valid_zip_codes)
    match_count = matches.sum()
    
    print(f"Found {match_count} matching ZIP codes")
    
    if match_count > 0:
        df_filtered = df[matches]  # Return original df with matching rows
        removed_count = original_count - len(df_filtered)
        print(f"   ZIP filter removed {removed_count} rows")
        return df_filtered, removed_count
    else:
        print(f"   No ZIP codes matched - returning original data")
        return df, 0

def filter_and_save(csv_file_path, filter_values, folder_path, client_name, fips_to_county_map, valid_zip_codes=None):
    """
    Enhanced version with improved ZIP filtering capability.
    """
    print(f"\nProcessing file: {os.path.basename(csv_file_path)}")
    
    # Load CSV file into a DataFrame
    try:
        df = pd.read_csv(csv_file_path)
        print(f"Loaded {len(df)} rows from CSV file")
    except Exception as e:
        print(f"Error loading CSV file: {e}")
        return

    # Debug: Print column names
    print("Columns in CSV:", df.columns.tolist())

    # Strip whitespace from column names
    df.columns = df.columns.str.strip()

    # Ensure the 'FIPS' column exists
    if 'FIPS' not in df.columns:
        raise ValueError("The column 'FIPS' does not exist in the provided CSV file.")

    # Convert 'FIPS' column to string and remove spaces if needed
    df['FIPS'] = df['FIPS'].astype(str).str.strip()

    print(f"\n--- Filtering Process ---")
    
    # 1. Filter rows where the 'FIPS' column contains any of the filter values
    print(f"1. Applying FIPS filter for codes: {filter_values}")
    initial_count = len(df)
    filtered_df = df[df['FIPS'].isin(filter_values)]
    fips_filtered_count = len(filtered_df)
    print(f"   After FIPS filter: {fips_filtered_count} rows (removed {initial_count - fips_filtered_count})")

    if len(filtered_df) == 0:
        print("Warning: No rows remaining after FIPS filtering!")
        return

    # 2. Apply ZIP code filtering if available
    if valid_zip_codes:
        print(f"2. Applying ZIP code filter ({len(valid_zip_codes)} valid ZIP codes)")
        filtered_df, zip_removed = filter_by_zip_codes(filtered_df, valid_zip_codes)
        zip_filtered_count = len(filtered_df)
        print(f"   After ZIP filter: {zip_filtered_count} rows (removed {zip_removed})")
        
        if len(filtered_df) == 0:
            print("Warning: No rows remaining after ZIP filtering!")
            return
    else:
        print("2. Skipping ZIP filter (no valid ZIP codes loaded)")
        zip_filtered_count = fips_filtered_count

    # 3. Replace 'FIPS' values with corresponding county names using the provided mapping
    print(f"3. Replacing FIPS codes with county names")
    filtered_df = filtered_df.copy()  # Avoid SettingWithCopyWarning
    filtered_df['FIPS'] = filtered_df['FIPS'].apply(lambda x: fips_to_county_map.get(x, x))

    # 4. Add a new column for the client name
    filtered_df['Client'] = client_name

    # 5. Format date columns to MM/YYYY if they exist
    date_columns_to_check = ['LAST SALE DATE', 'Sale Date', 'Last Sale Date']
    for date_col in date_columns_to_check:
        if date_col in filtered_df.columns:
            print(f"5. Formatting '{date_col}' column to MM/YYYY")
            try:
                filtered_df[date_col] = pd.to_datetime(filtered_df[date_col], errors='coerce').dt.strftime('%m/%Y')
                break  # Only format the first date column found
            except Exception as e:
                print(f"Error formatting '{date_col}': {e}")

    # 6. Rename columns to standardized names
    print(f"Current columns before rename: {filtered_df.columns.tolist()}")
    filtered_df = rename_columns(filtered_df)

    # Create folder if it doesn't exist
    os.makedirs(folder_path, exist_ok=True)

    # Generate filename with client name and timestamp
    from datetime import datetime
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    filename = f"market_filtered_data_{client_name.replace(' ', '_')}_{timestamp}.xlsx"
    excel_file_path = os.path.join(folder_path, filename)

    # Save the filtered DataFrame to the Excel file
    filtered_df.to_excel(excel_file_path, index=False)

    # Final summary
    print(f"\n--- Processing Summary ---")
    print(f"Original rows: {initial_count}")
    print(f"After FIPS filter: {fips_filtered_count}")
    if valid_zip_codes:
        print(f"After ZIP filter: {zip_filtered_count}")
    print(f"Final rows: {len(filtered_df)}")
    print(f"Final columns: {filtered_df.columns.tolist()}")
    print(f"Client: {client_name}")
    print(f"Output file: {excel_file_path}")

def process_multiple_files(input_folder, filter_values, output_folder, client_name, fips_to_county_map, valid_zip_codes=None):
    """
    Process multiple CSV files in a folder.
    """
    # Get all CSV files in the input folder
    csv_files = [f for f in os.listdir(input_folder) if f.endswith('.csv')]
    
    if not csv_files:
        print(f"No CSV files found in {input_folder}")
        return
    
    print(f"Found {len(csv_files)} CSV files to process:")
    for file in csv_files:
        print(f"  - {file}")
    
    # Process each file with progress bar
    for file in tqdm(csv_files, desc="Processing CSV files"):
        csv_file_path = os.path.join(input_folder, file)
        try:
            filter_and_save(csv_file_path, filter_values, output_folder, client_name, fips_to_county_map, valid_zip_codes)
        except Exception as e:
            print(f"Error processing {file}: {e}")
            continue
    
    print(f"\nAll files processed. Check output folder: {output_folder}")

# --- Main Execution ---

# Configuration
csv_file_path = r"C:\Users\LENOVO\Documents\8020\py_scripts\Benchmark_report\FIP_csv_filter\Input\market_deals (2).csv"
input_folder = r"C:\Users\LENOVO\Documents\8020\py_scripts\Benchmark_report\FIP_csv_filter\Input"  # For batch processing
filter_values = ['42101']  # FIPS codes to filter
output_folder = r"C:\Users\LENOVO\Documents\8020\py_scripts\Benchmark_report\FIP_csv_filter\Output"

# Define the mapping of FIPS codes to county names
fips_to_county_map = {
    '42101': 'Philadelphia',
}

print("FIPS to County Filter with ZIP Code Enhancement")
print("=" * 50)

# Get client name
client_name = input("Enter the client name: ")

# Load valid ZIP codes from external file
print(f"\nLoading ZIP codes from: {ZIP_FILTER_FILE_PATH}")
print(f"Looking for column: '{ZIP_COLUMN_NAME}'")
valid_zips = load_valid_zip_codes(ZIP_FILTER_FILE_PATH)

if valid_zips:
    print(f"ZIP filter is active with {len(valid_zips)} valid codes")
    # Show sample of loaded ZIP codes for verification
    sample_zips = list(valid_zips)[:10]
    print(f"Sample ZIP codes: {sample_zips}")
else:
    print("ZIP filter is inactive")

# Ask user if they want to process single file or multiple files
process_choice = input("\nProcess single file (s) or all files in folder (f)? [s/f]: ").lower().strip()

if process_choice == 'f':
    # Process all CSV files in the input folder
    process_multiple_files(input_folder, filter_values, output_folder, client_name, fips_to_county_map, valid_zips)
else:
    # Process single file
    filter_and_save(csv_file_path, filter_values, output_folder, client_name, fips_to_county_map, valid_zips)