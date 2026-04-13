import pandas as pd
import os
from tqdm import tqdm

# Define the set of valid options for easy lookup.
VALID_TYPES = {'DEAL', 'FULFILLMENT'}

# Path to ZIP codes filter file
ZIP_FILTER_PATH = r"C:\Users\LENOVO\Documents\8020\py_scripts\Benchmark_report\ZIP_benchmark.xlsx"

# --- Functions ---

def load_zip_filter(zip_filter_path):
    """Load ZIP codes from external file for filtering."""
    try:
        if not os.path.exists(zip_filter_path):
            print(f"ZIP filter file not found at: {zip_filter_path}")
            return []
        
        zip_df = pd.read_excel(zip_filter_path)
        
        if 'ZIP codes' not in zip_df.columns:
            print(f"Column 'ZIP codes' not found in ZIP filter file. Available columns: {list(zip_df.columns)}")
            return []
        
        zip_codes = zip_df['ZIP codes'].dropna().astype(str).tolist()
        print(f"Loaded {len(zip_codes)} ZIP codes for filtering.")
        return zip_codes
    except Exception as e:
        print(f"Could not load ZIP filter file. Error: {e}")
        return []

def determine_file_type(filename):
    """Determine file type based on filename structure."""
    filename_lower = filename.lower()
    if "_deals" in filename_lower:
        return "DEAL"
    elif "fulfillment" in filename_lower:
        return "FULFILLMENT"
    else:
        return None

def filter_unwanted_names(df, column_name, unwanted_names):
    """Filters rows from a DataFrame based on a list of unwanted names in a specific column."""
    pattern = '|'.join(unwanted_names)
    return df[~df[column_name].str.contains(pattern, case=False, na=False)]

def process_excel_files_v2(folder_path, output_folder, zip_codes_filter):
    """
    Processes Excel files in a folder. Automatically detects file type and applies appropriate filters.
    """
    processing_messages = []
    total_properties_processed = 0  # Initialize a counter for the total length
    files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    
    if not files:
        print("No Excel files (.xlsx) found in the specified folder.")
        return
    
    for file in tqdm(files, desc="Processing files"):
        file_path = os.path.join(folder_path, file)
        
        # Determine file type based on filename
        selected_type = determine_file_type(file)
        if selected_type is None:
            print(f"Skipping file {file} - could not determine type from filename.")
            continue
        
        print(f"\nProcessing {file} as {selected_type} type...")
        
        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            print(f"Could not read file {file}. Error: {e}")
            continue
        
        # --- Main processing logic ---
        
        # 1. Apply ZIP codes filter if provided
        if zip_codes_filter and 'ZIP' in df.columns:
            initial_count = len(df)
            df = df[df['ZIP'].astype(str).isin(zip_codes_filter)]
            filtered_count = len(df)
            print(f"ZIP filter applied: {initial_count} -> {filtered_count} records")
        
        # 2. Conditionally remove unwanted "OWNER FULL NAMES" only for FULFILLMENT type
        if selected_type == 'FULFILLMENT':
            unwanted_names = [
                "Given Not", "Record", "Available", "Bank ", "Church ", "School",
                "Cemetery", "Not given", "University", "College", "Owner",
                "Hospital", "County", "City of", "Not Provided Name"
            ]
            df = filter_unwanted_names(df, 'OWNER FULL NAME', unwanted_names)
        
        # 3. Remove duplicates based on column combinations (applied to all types)
        df.drop_duplicates(subset=['OWNER FULL NAME', 'ADDRESS', 'ZIP'], inplace=True)
        
        # 4. Get the length of the current DataFrame and add it to the total
        current_file_length = len(df)
        total_properties_processed += current_file_length
        print(f"Number of properties in {file} after cleaning: {current_file_length}")
        
        # 5. Save the modified file to the output folder (applied to all types)
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        
        output_file_path = os.path.join(output_folder, f"output - {file}")
        df.to_excel(output_file_path, index=False)
        
        processing_messages.append(f"File {file} ({selected_type}) processed and saved to {output_file_path}.")
    
    print("\n--- Processing Summary ---")
    for message in processing_messages:
        print(message)
    
    # Print the final total count
    print(f"\nTotal properties processed across all files: {total_properties_processed}")

# --- Script Execution ---

# Load ZIP codes filter
zip_codes_filter = load_zip_filter(ZIP_FILTER_PATH)

# Call the main function with ZIP filter.
process_excel_files_v2(
    folder_path=r"C:\Users\LENOVO\Documents\8020\py_scripts\Benchmark_report\Input file",
    output_folder=r"C:\Users\LENOVO\Documents\8020\py_scripts\Benchmark_report\Output file",
    zip_codes_filter=zip_codes_filter
)