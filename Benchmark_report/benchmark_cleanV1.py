import pandas as pd import os from tqdm import tqdm

# Define the set of valid options for easy lookup. VALID_TYPES = {'DEAL', 'FULFILLMENT', 'MARKET'}

# Renamed 'type' to 'file_type' to avoid conflict with the built-in Python function. file_type = None

while True: # Prompt the user for input and immediately convert it to uppercase. user_input = input("Enter the file type (DEAL, FULFILLMENT, MARKET): ").upper()

# Check if the capitalized input is in our set of valid types. if user_input in VALID_TYPES: # If the input is valid, store it and break the loop. file_type = user_input print(f"File type selected: {file_type}") break else: # If the input is invalid, inform the user and the loop will continue. print("Invalid input. Please enter 'DEAL', 'FULFILLMENT', or 'MARKET'.")

# --- Functions ---

def filter_unwanted_names(df, column_name, unwanted_names): """Filters rows from a DataFrame based on a list of unwanted names in a specific column.""" pattern = '|'.join(unwanted_names) return df[~df[column_name].str.contains(pattern, case=False, na=False)]

def process_excel_files_v2(folder_path, output_folder, selected_type): """ Processes Excel files in a folder. It filters unwanted names only for 'FULFILLMENT' type, but removes duplicates for all types. """ processing_messages = [] total_properties_processed = 0 # Initialize a counter for the total length files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

if not files: print("No Excel files (.xlsx) found in the specified folder.") return

for file in tqdm(files, desc="Processing files"): file_path = os.path.join(folder_path, file)

try: df = pd.read_excel(file_path) except Exception as e: print(f"Could not read file {file}. Error: {e}") continue

# --- Main processing logic ---

# 1. Conditionally remove unwanted "OWNER FULL NAMES" only for FULFILLMENT type if selected_type == 'FULFILLMENT': unwanted_names = [ "Given Not", "Record", "Available", "Bank ", "Church ", "School", "Cemetery", "Not given", "University", "College", "Owner", "Hospital", "County", "City of", "Not Provided Name" ] df = filter_unwanted_names(df, 'OWNER FULL NAME', unwanted_names)

# 2. Remove duplicates based on column combinations (applied to all types) df.drop_duplicates(subset=['OWNER FULL NAME', 'ADDRESS', 'ZIP'], inplace=True)

# 3. Get the length of the current DataFrame and add it to the total current_file_length = len(df) total_properties_processed += current_file_length print(f"\nNumber of properties in {file} after cleaning: {current_file_length}")

# 4. Save the modified file to the output folder (applied to all types) if not os.path.exists(output_folder): os.makedirs(output_folder)

output_file_path = os.path.join(output_folder, f"output - {file}") df.to_excel(output_file_path, index=False)

processing_messages.append(f"File {file} processed and saved to {output_file_path}.")

print("\n--- Processing Summary ---") for message in processing_messages: print(message)

# Print the final total count print(f"\nTotal properties processed across all files: {total_properties_processed}")

# --- Script Execution ---

# Call the main function with the user-selected file_type. process_excel_files_v2( folder_path=r"C:\Users\LENOVO\Documents\8020\AutomatedManualFulfillment\Income File", output_folder=r"C:\Users\LENOVO\Documents\8020\AutomatedManualFulfillment\Output file", selected_type=file_type )