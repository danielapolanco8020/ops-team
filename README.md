# Clean Name Processor for Real Estate Data

## Overview
This Python script is designed to process Excel files containing real estate data, specifically to clean and split full names of property owners into first and last names. It's particularly useful for real estate investors and analysts looking to normalize and analyze their datasets for outbound marketing or data analysis purposes.

## Features
- Processes multiple Excel files in a specified directory.
- Splits "OWNER FULL NAME" into "OWNER FIRST NAME" and "OWNER LAST NAME" based on predefined lists of common names and a set of heuristics for identifying business entities.
- Identifies and handles business entities differently to maintain data integrity.
- Outputs the processed files to a specified output directory, preserving the original data structure with the updated name columns.

## Installation Guide

### Prerequisites
- Python 3.6 or higher
- pandas
- tqdm
- OpenPyXL

### Steps
1. **Clone the Repository**
   ```sh
   git clone https://yourrepositorylink.com/path/to/CleanNameProcessor.git
   cd CleanNameProcessor
   ```

2. **Set Up a Virtual Environment** (optional but recommended)
   ```sh
   python3 -m venv venv
   source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
   ```

3. **Install Dependencies**
   ```sh
   pip install pandas tqdm openpyxl
   ```

4. **Configuration**
   Before running the script, ensure you have `vars.py` configured with your specific paths and name lists:
   - `data_folder`: Directory containing the Excel files to process.
   - `output_folder`: Directory where the processed files will be saved.
   - `common_first_names`, `common_last_names`: Lists of common first and last names to aid in name splitting.

## How It Works
1. **Prepare Your Excel Files**: Place all your Excel files in the `data_folder`. Each file should have an "OWNER FULL NAME" column.

2. **Run the Script**
   ```sh
   python3 main.py
   ```
   The script automatically processes each Excel file in the `data_folder`, splitting full names into first and last names and saving the modified files to the `output_folder`.

3. **Process Flow**
   - The script iterates over each Excel file, checking for the required columns.
   - It applies a set of heuristics and predefined lists to accurately split full names into first and last names.
   - Business entities are detected and handled differently to ensure data accuracy.
   - Progress is displayed in the terminal via a progress bar.

4. **Output**
   Processed files are saved in the specified `output_folder` with names intact, including the updated name columns.

## Notes
- The script is designed to handle common naming conventions and business entity formats but may require adjustments for specific datasets or unique cases.
- Always review the output for accuracy and make any necessary manual corrections.

## Support
For issues, suggestions, or contributions, please contact [your contact information] or submit a pull request.
