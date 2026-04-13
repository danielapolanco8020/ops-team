import os
import pandas as pd
import re
import matplotlib.pyplot as plt
from datetime import datetime
from tqdm import tqdm

def extract_info_from_filename(filename):
    """Extracts date, quarter, and marketing channel from the filename."""
    date_match = re.search(r'\d{4}-\d{2}-\d{2}', filename)
    channel_match = re.search(r'(Cold Calling|Sms|Direct Mail)', filename, re.IGNORECASE)
    
    if not date_match or not channel_match:
        return None, None, None
    
    date_str = date_match.group()
    channel = channel_match.group().title()
    
    date_obj = datetime.strptime(date_str, '%Y-%m-%d')
    month = date_obj.month
    quarter = f'Q{((date_obj.month - 1) // 3) + 1}'
    
    return month, quarter, channel

def process_files(input_folder, output_file):
    """Processes multiple Excel files and generates a summary."""
    summary_data = []
    files = [f for f in os.listdir(input_folder) if f.endswith(".xlsx") or f.endswith(".xls")]
    
    for file in tqdm(files, desc="Processing Files", unit="file"):
        file_path = os.path.join(input_folder, file)
        df = pd.read_excel(file_path)
        
        count_columns = [col for col in df.columns if ' count' in col.lower()]
        if count_columns:
            zero_rows = df[count_columns].eq(0).any(axis=1).sum()
            month, quarter, channel = extract_info_from_filename(file)
            if channel:
                summary_data.append({'Month': month, 'Quarter': quarter, 'Channel': channel, 'Zero Rows': zero_rows})
    
    summary_df = pd.DataFrame(summary_data)
    summary_df.to_excel(output_file, index=False)
    
    plot_bar_chart(summary_df)

def plot_bar_chart(df):
    """Generates a bar chart based on the summarized data."""
    pivot_table = df.pivot_table(index='Quarter', columns='Channel', values='Zero Rows', aggfunc='sum', fill_value=0)
    
    pivot_table.plot(kind='bar', figsize=(10, 6))
    plt.xlabel("Quarter")
    plt.ylabel("Quantity of Zero Rows")
    plt.title("Zero Count Rows by Marketing Channel and Quarter")
    plt.xticks(rotation=0)
    plt.legend(title="Marketing Channel")
    plt.tight_layout()
    plt.show()

# Example usage
input_folder = r"C:\Users\LENOVO\Documents\8020\py_scripts\new_property_count\input"
output_file = "output_summary.xlsx"
process_files(input_folder, output_file)
