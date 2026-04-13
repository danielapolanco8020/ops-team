import os
import pandas as pd
import re
from datetime import datetime

def check_duplicates_by_columns(df, columns):
    """Check for duplicates based on specified columns."""
    duplicate_rows = df[df.duplicated(subset=columns, keep=False)]
    return duplicate_rows

def check_repeated_values(df, column, compare_columns):
    """Check for repeated values in a column with different values in other columns."""
    repeated_values = df[df.duplicated(subset=[column], keep=False)]
    different_values = repeated_values[repeated_values.duplicated(subset=compare_columns, keep=False)]
    return different_values

def audit(file_path):
    """Perform comprehensive audit on property data Excel file."""
    # Load Excel file
    df = pd.read_excel(file_path)

    # Task 0: Print the number of properties
    actual_properties = len(df)
    print(f"Number of properties in file: {actual_properties}")

    results = {
        "Duplicates review #1": [],
        "Duplicates review #2": [],
        "Unique Folio": [],
        "Owner Full Name value review": [],
        "Owner Full Name complete": [],
        "Owner Last Name complete": [],
        "Address complete": [],
        "Zip complete": [],
        "County review": [],
        "Mailing Address complete": [],
        "Mailing Zip complete": [],
        "Action Plans review": [],
        "Score greater than 0 review": [],
        "Buybox Score greater than 0 review": [],
        "Property type review": [],
        "Tags review": [],
        "Absentee address review": [],
        "Urgent score": [],
        "High score": []
    }

    # Task 16: Check for duplicates in "MAILING ADDRESS" and "MAILING ZIP" columns
    duplicates_1 = check_duplicates_by_columns(df, ["MAILING ADDRESS", "MAILING ZIP"])
    if not duplicates_1.empty:
        results["Duplicates review #1"] = duplicates_1
        print("Duplicates review #1: Failed")
        print(f"Number of Duplicates Mailing Zip and Address: {len(duplicates_1)}")
    else:
        print("Duplicates review #1: Pass")

    # Task 17: Check for duplicates in "OWNER FULL NAME", "ADDRESS", and "ZIP" columns
    duplicates_2 = check_duplicates_by_columns(df, ["OWNER FULL NAME", "ADDRESS", "ZIP"])
    if not duplicates_2.empty:
        results["Duplicates review #2"] = duplicates_2
        print("Duplicates review #2: Failed")
        print(f"Number of Duplicates Owner Full Name, Address and Zip: {len(duplicates_2)}")
    else:
        print("Duplicates review #2: Pass")

    # Task 1: Check for repeated values in "FOLIO" column and compare specified columns
    repeated_values = check_repeated_values(df, "FOLIO", ["OWNER FULL NAME", "ADDRESS", "ZIP"])
    if not repeated_values.empty:
        results["Unique Folio"] = repeated_values
        print("Unique Folio: Failed")
        print(f"Number of Duplicates Folio with different info: {len(repeated_values)}")
    else:
        print("Unique Folio: Pass")

    # Task 2: Review values in "OWNER FULL NAME" column
    keywords = ["Given ", "Not ", "Record ", "Available ", "Bank ", "Church ", "School ", "Cemetery ", 
                "Not given ", "University", "College", "Owner ", "Hospital ", "County ", "City of", 
                "Unknown ", "Not Provided "]
    owner_full_name_review = df[df["OWNER FULL NAME"].str.contains("|".join(keywords), na=False)]
    invalid_owner_names = owner_full_name_review[~owner_full_name_review["OWNER FULL NAME"].isin(keywords)]

    if not invalid_owner_names.empty:
        results["Owner Full Name value review"] = owner_full_name_review
        print("Owner Full Name value review: Failed")
        print(f"Number of Properties with invalid Owner Name: {len(invalid_owner_names)}")
        print("Invalid Owner Names:")
        print(invalid_owner_names["OWNER FULL NAME"].tolist())
    else:
        print("Owner Full Name value review: Pass")

    # Task 3: Check for empty cells in "OWNER FULL NAME" column
    if df["OWNER FULL NAME"].isnull().any():
        results["Owner Full Name complete"].append("Empty cells")
        print("Owner Full Name complete: Failed")
    else:
        print("Owner Full Name complete: Pass")

    # Task 4: Check for empty cells in "OWNER LAST NAME" column
    if df["OWNER LAST NAME"].isnull().any():
        results["Owner Last Name complete"].append("Empty cells")
        print("Owner Last Name complete: Failed")
    else:
        print("Owner Last Name complete: Pass")

    # Task 5: Check for empty cells in "ADDRESS" column
    if df["ADDRESS"].isnull().any():
        results["Address complete"].append("Empty cells")
        print("Address complete: Failed")
    else:
        print("Address complete: Pass")

    # Task 6: Check for empty cells in "ZIP" column
    if df["ZIP"].isnull().any():
        results["Zip complete"].append("Empty cells")
        print("Zip complete: Failed")
    else:
        print("Zip complete: Pass")

    # Task 7: Print unique values in "COUNTY" column
    unique_counties = df["COUNTY"].dropna().str.lower().unique()
    print("Unique Counties:")
    print(unique_counties)
    results["County review"] = df  # Store all data as we can't validate without user input

    # Task 8: Check for empty cells in "MAILING ADDRESS" column
    if df["MAILING ADDRESS"].isnull().any():
        results["Mailing Address complete"].append("Empty cells")
        print("Mailing Address complete: Failed")
    else:
        print("Mailing Address complete: Pass")

    # Task 9: Check for empty cells in "MAILING ZIP" column
    if df["MAILING ZIP"].isnull().any():
        results["Mailing Zip complete"].append("Empty cells")
        print("Mailing Zip complete: Failed")
    else:
        print("Mailing Zip complete: Pass")

    # Task 10: Review values in "ACTION PLANS" column - Urgent score
    action_plans_review = df[df["ACTION PLANS"] == "30 DAYS"]
    if not action_plans_review.empty:
        urgent_score_review = action_plans_review[df["SCORE"] < 746]
        if not urgent_score_review.empty:
            results["Urgent score"] = urgent_score_review
            print("Urgent score: Fail")
            print(f"Number of Urgent with score below 746: {len(urgent_score_review)}")
        else:
            print("Urgent score: Pass")
    else:
        print("Urgent score: Pass")

    # Task 11: Review values in "ACTION PLANS" column - High score
    high_score_review = df[(df["ACTION PLANS"] == "60 DAYS") & (df["SCORE"] < 545)]
    if not high_score_review.empty:
        results["High score"] = high_score_review
        print("High score: Fail")
        print(f"Number of High with score below 545: {len(high_score_review)}")
    else:
        print("High score: Pass")

    # Task 12: Review values in "PROPERTY STATUS" column
    property_status_review = df[df["PROPERTY STATUS"].isnull()]
    if not property_status_review.empty:
        results["Property status review"] = property_status_review
        print("Property status review: Failed")
    else:
        print("Property status review: Pass")

    # Task 13: Print unique values in "PROPERTY TYPE" column
    unique_property_types = df["PROPERTY TYPE"].dropna().str.lower().unique()
    print("Unique Property Types:")
    print(unique_property_types)
    results["Property type review"] = df  # Store all data as we can't validate without user input

    # Task 14: Review values in "TAGS" column
    tags_list = ["Liti", "DNC", "donotmail", "Takeoff", "Undeli", "Return", "Dead", 
                 "Do Not Mail", "Dono", "Do no", "Available"]
    df["TAGS"] = df["TAGS"].astype(str).fillna("")
    tags_review = df[df["TAGS"].str.contains("|".join(tags_list), na=False)]
    if not tags_review.empty:
        results["Tags review"] = tags_review
        print("Tags review: Failed")
        print(f"Number of properties with unwanted Tags: {len(tags_review)}")
    else:
        print("Tags review: Pass")

    # Task 15: Review values in "ABSENTEE" column
    absentee_review = df[(df["ABSENTEE"] >= 1) & (df["ADDRESS"] == df["MAILING ADDRESS"])]
    if not absentee_review.empty:
        results["Absentee address review"] = absentee_review
        print("Absentee address review: Failed")
        print(f"Number of absentee properties with same Address and Mailing Address: {len(absentee_review)}")
    else:
        print("Absentee address review: Pass")

    # Task 20: Check for "Void", "Null", "Failed" in "PHONE TYPE" if file name includes "Cold Calling"
    if "cold calling" in file_path.lower() or "cc" in file_path.lower():
        phone_type_columns = [col for col in df.columns if col.startswith("PHONE TYPE")]
        for col in phone_type_columns:
            try:
                phone_type_check = df[df[col].str.contains("void|null|failed", case=False, na=False)]
            except AttributeError:
                df[col] = df[col].astype(str)
                phone_type_check = df[df[col].str.contains("void|null|failed", case=False, na=False)]

            if not phone_type_check.empty:
                results[f"{col} check"] = phone_type_check
                print(f"Number of properties with wrong phone number type in {col}: {len(phone_type_check)}")
                print(f"{col} check: Fail")
            else:
                print(f"{col} check: Pass")

    # Task 21: Check for "Void", "Null", "Failed", "Landline" in "PHONE TYPE" if file name includes "Sms"
    if "sms" in file_path.lower():
        phone_type_columns = [col for col in df.columns if col.startswith("PHONE TYPE")]
        for col in phone_type_columns:
            try:
                phone_type_check = df[df[col].str.contains("void|null|failed|landline", case=False, na=False)]
            except AttributeError:
                df[col] = df[col].astype(str)
                phone_type_check = df[df[col].str.contains("void|null|failed|landline", case=False, na=False)]

            if not phone_type_check.empty:
                results[f"{col} check"] = phone_type_check
                print(f"Number of properties with wrong phone number type in {col}: {len(phone_type_check)}")
                print(f"{col} check: Fail")
            else:
                print(f"{col} check: Pass")

    return results

def phone_count_check(file_path):
    """Check phone number counts for SMS or Cold Calling files."""
    # Load Excel file
    df = pd.read_excel(file_path)

    # Check if the file name contains keywords related to phone type or SMS
    keywords = ['sms', 'cc', 'cold calling']

    # Extract the file name from the file path
    file_name = os.path.basename(file_path)

    # Check if any keyword is present in the file name using regular expressions
    if any(re.search(r'\b{}\b'.format(re.escape(keyword)), file_name, re.IGNORECASE) for keyword in keywords):
        print("\n--- PHONE NUMBER COUNT ---")
        print(f'File Analyzed: {file_path}')

        # Count Properties with Data in Phone Type or Phone Number Columns
        properties_with_data = df[(df.filter(like='PHONE TYPE').notna().any(axis=1)) | 
                                 (df.filter(like='PHONE NUMBER').notna().any(axis=1))]
        properties_with_data_count = len(properties_with_data)
        print(f"Properties with Active Phone Numbers: {properties_with_data_count}")

        # Count Empty Rows in Phone Type or Phone Number Columns
        empty_rows = df[(df.filter(like='PHONE TYPE').isna().all(axis=1)) & 
                       (df.filter(like='PHONE NUMBER').isna().all(axis=1))]
        empty_rows_count = len(empty_rows)
        print(f"Properties without Active Phone Numbers: {empty_rows_count}")

        # Count Rows without "SKIPTRACE" in TAGS column
        skiptrace_count = len(empty_rows[~empty_rows["TAGS"].str.contains("skip", na=False, case=False)])
        print(f"Properties without 'SKIPTRACE': {skiptrace_count}")
    else:
        print("File name does not contain keywords related to SMS or Cold Call.")

def main():
    """Main function to run the audit and phone count check for all Excel files in a folder."""
    print("=" * 60)
    print("PROPERTY DATA AUDITOR")
    print("=" * 60)
    
    # Get folder path from user
    folder_path = r"C:\Users\LENOVO\Documents\8020\py_scripts\auditV2\Input"
    
    # Remove quotes if present
    folder_path = folder_path.strip('"').strip("'")
    
    # Check if folder exists
    if not os.path.isdir(folder_path):
        print(f"Error: Folder not found at {folder_path}")
        return
    
    # Get list of Excel files in the folder
    excel_files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.xlsx', '.xls'))]
    
    if not excel_files:
        print("Error: No Excel files (.xlsx or .xls) found in the folder")
        return
    
    print(f"\nFound {len(excel_files)} Excel file(s) in the folder:")
    for file in excel_files:
        print(f"- {file}")
    
    # Process each Excel file
    for file in excel_files:
        file_path = os.path.join(folder_path, file)
        print(f"\nAnalyzing file: {file}")
        print(f"Full path: {file_path}")
        
        # Run audit
        print("\n" + "=" * 60)
        print(f"STARTING AUDIT FOR {file}")
        print("=" * 60)
        results = audit(file_path)
        
        if results is None:
            print(f"\nAudit was cancelled for {file}.")
            continue
        
        # Run phone count check
        print("\n" + "=" * 60)
        phone_count_check(file_path)
        print("=" * 60)
        
        print(f"\nAUDIT COMPLETE FOR {file}")
        print("=" * 60)
    
    print("\n" + "=" * 60)
    print("ALL FILES PROCESSED")
    print("=" * 60)

if __name__ == "__main__":
    main()