import pandas as pd
import re

# Define the unwanted_names list (matching the original Python code)
unwanted_names = [
    "Given Not", "Record", "Available", "Bank ", "Church ",
    "School", "Cemetery", "Not given", "University", "College",
    "Owner", "Hospital", "County", "City of", "Not Provided Name"
]

# Function to extract the first matching unwanted name (mimics Excel formula)
def find_unwanted_name(text, unwanted_names):
    """
    Returns the first unwanted_names string that appears in text (trimmed).
    Case-insensitive, returns empty string if no match.
    """
    if pd.isna(text) or not isinstance(text, str):
        return ""
    text_lower = text.lower()
    for name in unwanted_names:
        if name.strip().lower() in text_lower:
            return name.strip()  # Return trimmed name (e.g., "Church")
    return ""

# Function to extract the full word containing the unwanted name
def extract_full_word(text, unwanted_names):
    """
    Returns the full word from text that contains an unwanted_names string.
    Uses regex for word boundaries, case-insensitive.
    """
    if pd.isna(text) or not isinstance(text, str):
        return ""
    text_str = text
    text_lower = text_str.lower()
    for name in unwanted_names:
        # Create regex pattern: \b\w*name\w*\b for whole word
        pattern = r'\b\w*' + re.escape(name.strip().lower()) + r'\w*\b'
        match = re.search(pattern, text_lower)
        if match:
            return text_str[match.start():match.end()]  # Return full word (e.g., "Churchill")
    return ""

# Function to process the Rejected_Properties Excel file
def process_rejected_properties(input_file= r"C:\Users\LENOVO\Documents\8020\AutomatedManualFulfillment\Output file\Rejected_Properties.xlsx", 
                               output_file='Rejected_Properties_Output.xlsx', 
                               sheet_name='Sheet1', 
                               column_name='OWNER FULL NAME'):
    """
    Process Rejected_Properties.xlsx to extract unwanted names and full words.
    Args:
        input_file (str): Path to input Excel file
        output_file (str): Path to save output Excel file
        sheet_name (str): Name of the sheet to read
        column_name (str): Name of column containing owner names
    """
    try:
        # Read the Excel file
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        
        # Ensure the column exists
        if column_name not in df.columns:
            raise ValueError(f"Column '{column_name}' not found in {input_file}, sheet {sheet_name}. Available columns: {list(df.columns)}")
        
        # Apply the functions
        df['Unwanted_Name'] = df[column_name].apply(lambda x: find_unwanted_name(x, unwanted_names))
        df['Full_Word'] = df[column_name].apply(lambda x: extract_full_word(x, unwanted_names))
        
        # Save to output Excel file
        df.to_excel(output_file, index=False)
        print(f"Results saved to {output_file}")
        print("\nSample output (first 5 rows):")
        print(df[[column_name, 'Unwanted_Name', 'Full_Word']].head())
        
    except FileNotFoundError:
        print(f"Error: Input file {input_file} not found. Please ensure the file exists in the script's directory.")
    except ValueError as ve:
        print(f"Error: {str(ve)}")
    except Exception as e:
        print(f"Error: {str(e)}. Please check the file format, sheet name, or column name.")

# Example with sample data (for testing if file is unavailable)
def run_sample():
    """
    Run with sample data if the Excel file is not available.
    """
    # Sample DataFrame
    data = pd.DataFrame({
        'OWNER FULL NAME': [
            "Trinity United Methodist Church Board Of",
            "Churchill",
            "Bank of America",
            "Givens Estate",
            "John Doe",
            "",
            "Banks Family Trust"
        ]
    })
    
    # Apply functions
    data['Unwanted_Name'] = data['OWNER FULL NAME'].apply(lambda x: find_unwanted_name(x, unwanted_names))
    data['Full_Word'] = data['OWNER FULL NAME'].apply(lambda x: extract_full_word(x, unwanted_names))
    
    # Save to output
    output_file = 'Rejected_Properties_Sample_Output.xlsx'
    data.to_excel(output_file, index=False)
    print(f"Sample results saved to {output_file}")
    print("\nSample output:")
    print(data)

if __name__ == "__main__":
    # Process the Rejected_Properties.xlsx file
    process_rejected_properties()
    
    # Uncomment to test with sample data if the file is unavailable
    # run_sample()