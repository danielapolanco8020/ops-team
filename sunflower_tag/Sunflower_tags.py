import pandas as pd
import os
from datetime import datetime
from dateutil.relativedelta import relativedelta

def batch_process_properties_strict():
    # --- Configuration ---
    input_folder = r"C:\Users\LENOVO\Documents\8020\py_scripts\Sunflower_tags\input"
    output_folder = r"C:\Users\LENOVO\Documents\8020\py_scripts\Sunflower_tags\output"
    
    # Create folders if they don't exist
    if not os.path.exists(input_folder):
        os.makedirs(input_folder)
        print(f"Created folder: '{input_folder}'. Please put your Excel files here.")
        return

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Get list of files
    files = [f for f in os.listdir(input_folder) if f.endswith('.xlsx') or f.endswith('.xls')]
    
    if not files:
        print(f"No Excel files found in '{input_folder}'.")
        return

    # --- Date Setup ---
    current_date = datetime.now()
    # 6 months ago from today
    cutoff_date = current_date - relativedelta(months=6)

    print(f"--- Analysis Settings ---")
    print(f"Today: {current_date.strftime('%Y-%m-%d')}")
    print(f"Cutoff: {cutoff_date.strftime('%Y-%m-%d')}")
    print("Logic: If ANY Skiptrace tag is recent (newer than cutoff), the property is ACTIVE.\n")

    # --- Logic Function ---
    def determine_status(cell_value):
        if not isinstance(cell_value, str):
            return 'Active' # Empty cells are Active/Ignored
            
        tags = [t.strip() for t in cell_value.split(',')]
        
        has_recent_tag = False
        has_old_tag = False
        
        for tag in tags:
            if tag.startswith("Skiptrace"):
                try:
                    date_part = tag.replace("Skiptrace", "").strip()
                    tag_date = datetime.strptime(date_part, "%B%Y")
                    
                    if tag_date >= cutoff_date:
                        # Found a recent tag! This "saves" the property immediately.
                        has_recent_tag = True
                        break # No need to check other tags
                    else:
                        # Found an old tag, but we keep checking in case there is a new one later
                        has_old_tag = True
                        
                except ValueError:
                    continue

        # --- Final Decision ---
        if has_recent_tag:
            return 'Active'
        elif has_old_tag:
            return 'OLDER_THAN_6_MONTHS'
        else:
            return 'Active' # No Skiptrace tags found at all

    # --- Processing Loop ---
    for filename in files:
        input_path = os.path.join(input_folder, filename)
        output_path = os.path.join(output_folder, f"Processed_{filename}")
        
        try:
            print(f"Processing: {filename}...", end=" ")
            df = pd.read_excel(input_path)
            
            if 'TAGS' not in df.columns:
                print(f"[SKIPPED] - Column 'TAGS' not found.")
                continue

            # Apply the new logic to create the column
            df['Tag_Analysis'] = df['TAGS'].apply(determine_status)
            
            # Count how many were actually flagged as old
            count_old = len(df[df['Tag_Analysis'] == 'OLDER_THAN_6_MONTHS'])
            
            df.to_excel(output_path, index=False)
            print(f"[DONE] - Identified {count_old} properties to update.")
            
        except Exception as e:
            print(f"[ERROR] - {e}")

    print(f"\nAll files processed in '{output_folder}'.")

# Run it
batch_process_properties_strict()