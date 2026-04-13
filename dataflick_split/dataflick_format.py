import pandas as pd
import os

def process_files(input_folder, output_folder, chunk_size=20000):
    # List all files in the input folder
    files = os.listdir(input_folder)

    # Loop through each file in the folder
    for file_name in files:
        if file_name.endswith('.xlsx'):  # Assuming all files are Excel files
            file_path = os.path.join(input_folder, file_name)

            # Read the Excel file
            df = pd.read_excel(file_path)

            # Keep only the desired columns
            columns_to_keep = ['property_address', 'property_city', 'property_state', 'property_zip',
                               'mailing_address', 'mailing_address_city', 'mailing_address_state', 'mailing_address_zip',
                               'phone_1', 'phone_1_type','phone_2', 'phone_2_type',
                               'phone_3', 'phone_3_type','phone_4', 'phone_4_type']
            df_subset = df[columns_to_keep]

            # Split large files into chunks
            if len(df_subset) > chunk_size:
                num_chunks = (len(df_subset) - 1) // chunk_size + 1
                for i in range(num_chunks):
                    chunk_start = i * chunk_size
                    chunk_end = min((i + 1) * chunk_size, len(df_subset))
                    df_chunk = df_subset.iloc[chunk_start:chunk_end]

                    # Write the chunk to a new Excel file
                    output_file_path = os.path.join(output_folder, f'{file_name.split(".")[0]}_chunk_{i+1}.xlsx')
                    df_chunk.to_excel(output_file_path, index=False)
            else:
                # Write the entire subset to a new Excel file
                output_file_path = os.path.join(output_folder, f'{file_name.split(".")[0]}_output.xlsx')
                df_subset.to_excel(output_file_path, index=False)

if __name__ == "__main__":
    # Input folder path
    input_folder_path = r"C:\Users\Roque Navas\Documents\8020\input_dataflick"

    # Output folder path
    output_folder_path = r"C:\Users\Roque Navas\Documents\8020\output_dataflick"

    # Create the output folder if it doesn't exist
    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)

    # Process files
    process_files(input_folder_path, output_folder_path)
    print("File Formatted")
