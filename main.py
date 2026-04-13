from vars import *
import pandas as pd
import os, re
from pathlib import Path
from tqdm import tqdm  # Import tqdm for the progress bar

def process_excel_files(data_folder, output_folder, common_first_names, common_last_names):
    # Ensure output directory exists
    Path(output_folder).mkdir(parents=True, exist_ok=True)

    excel_files = [file for file in os.listdir(data_folder) if file.endswith(".xlsx")]
    
    for file in tqdm(excel_files, desc="Processing Excel files"):
        file_path = os.path.join(data_folder, file)
        df = pd.read_excel(file_path)

        if "OWNER FULL NAME" in df.columns and "OWNER FIRST NAME" in df.columns and "OWNER LAST NAME" in df.columns:
            # Update OWNER FIRST NAME and OWNER LAST NAME columns
            names = df["OWNER FULL NAME"].apply(lambda full_name: clean_and_split_name(full_name, common_first_names, common_last_names))
            df["OWNER FIRST NAME"], df["OWNER LAST NAME"] = zip(*names)

            output_path = os.path.join(output_folder, file)
            df.to_excel(output_path, index=False)
        else:
            print(f"Required columns are missing in {file}")

import re

def clean_and_split_name(full_name, common_first_names, common_last_names):
    business_entities = ["trust", "investment", "properties", "prop", "capital", "acquisitions", "association", "inc", "incorporated", "and", "council", "rental"]
    name_terms = ["jr", "sr", "tr", "ii", "iii", "iv", "esq", "phd", "md", "aka", "fka", "tod", "dba", "mba", "cpa", "-estate of"]
    full_name = str(full_name).strip(" ,-.")

    # Split and lower case for processing
    name_words = full_name.lower().split()

    # Check and handle if 'llc' is in the name separately
    if 'llc' in name_words:
        # Identify and remove 'llc', set it as last_name
        name_words.remove('llc')
        first_name = ' '.join(name_words).title()
        last_name = 'LLC'
    elif any(entity in name_words for entity in business_entities):
        # Handle other business entities
        first_name, last_name = full_name, full_name  
        cleaned_full_name = full_name.lower().title()  # Example specific adjustment
    else:
        # Handle individual names
        words = full_name.split()
        if 'tr' in [word.lower() for word in words]:
            last_name = words[0]
            first_name = ' '.join(words[1:-1])  
        else:
            cleaned_name = re.sub(r'[^\w\s]', '', full_name)
            name_parts = cleaned_name.split()
            essential_parts = [word for word in name_parts if len(word) > 1 or (len(word) == 1 and len(name_parts) == 2)]
            if len(essential_parts) == 1:
                first_name, last_name = essential_parts[0], ''
            elif essential_parts:
                first_name, last_name = essential_parts[0], ' '.join(essential_parts[1:])
            else:
                first_name, last_name = name_parts[0], '' if name_parts else ''

        # Further cleaning for name correctness
        first_name_parts = first_name.split()
        last_name_parts = last_name.split()
        first_name = ' '.join(word for word in first_name_parts if word.lower() not in name_terms and len(word) > 1)
        last_name = ' '.join(word for word in last_name_parts if word.lower() not in name_terms and len(word) > 1)

        # Check against common names lists for order correction
        first_in_first_names = first_name.title() in common_first_names
        first_in_last_names = first_name.title() in common_last_names
        last_in_first_names = last_name.title() in common_first_names
        last_in_last_names = last_name.title() in common_last_names

        # Swap names under specific conditions
        if (first_in_last_names and not first_in_first_names and not last_in_last_names) or \
           (last_in_first_names and not last_in_last_names and not first_in_first_names):
            first_name, last_name = last_name, first_name

    return first_name.strip(), last_name.strip()


if __name__ == '__main__':
    process_excel_files(data_folder, output_folder, common_first_names, common_last_names)