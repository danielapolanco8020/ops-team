import pandas as pd
import numpy as np
from pathlib import Path
import openpyxl
import warnings
import os

# Hardcoded folder paths
INPUT_FOLDER = r"C:\Users\LENOVO\Documents\8020\py_scripts\Split_code\input"
OUTPUT_FOLDER = r"C:\Users\LENOVO\Documents\8020\py_scripts\Split_code\output"

# Function to validate required columns
def validate_columns(df, required_cols, client_name):
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Client {client_name}: Missing required columns: {', '.join(missing_cols)}")

# Function to print summary of process for all clients before selection
def print_process_summary():
    # MODIFIED: Added summary for Client 6
    summaries = {
        '1': "Client 1: Split all records into two templates (A and B) with balanced 30 DAYS/60 DAYS/90 DAYS plans, using user-specified proportions (default 50/50), using all rows.",
        '2': "Client 2: Split all records into two groups based on odd and even row positions in the input list.",
        '3': "Client 3: Include top X properties (by SCORE) in both templates, split remaining records into two templates (A and B) with balanced 30 DAYS/60 DAYS/90 DAYS plans, maintaining original row order.",
        '4': "Client 4: Split all records by user-specified PROPERTY TYPE values (case-sensitive) or grouped PROPERTY TYPEs (e.g., Group 1: Residential, Condo).",
        '5': "Client 5: Split all records by user-specified counties or grouped counties.",
        '6': "Client 6: Split all records into separate templates based on ACTION PLANS containing '30 DAYS', '60 DAYS', or '90 DAYS'.",
        'all': "All: Process splits for all clients: Client 1 (proportions), Client 2 (odd/even), Client 3 (top X), Client 4 (property type), Client 5 (county), and Client 6 (ACTION PLANS)."
    }
    print("\nProcess Summaries for All Clients:")
    for client, description in summaries.items():
        print(f"{description}")
    print()

# Function to prompt for Client 3 top X properties
def prompt_client_3_params(df):
    max_records = len(df)
    print(f"Total records available: {max_records}")
    while True:
        try:
            x = int(input(f"Enter number of top properties to preserve for Client 3 (1 to {max_records}): ").strip())
            if 1 <= x <= max_records:
                break
            print(f"Please enter a number between 1 and {max_records}.")
        except ValueError:
            print("Please enter a valid integer.")
    
    return x

# Function to prompt for Client 4 PROPERTY TYPE values or grouped PROPERTY TYPEs
def prompt_client_4_params(df):
    valid_prop_types = sorted(df['PROPERTY TYPE'].dropna().unique())
    if not valid_prop_types:
        raise ValueError("Client 4: No valid PROPERTY TYPE values found in data")
    print(f"Available PROPERTY TYPE values (case-sensitive): {', '.join(valid_prop_types)}")
    
    # Prompt for grouping option
    group_prop_types = input("Do you want to split Client 4 by grouped PROPERTY TYPEs? (yes/no): ").strip().lower() == 'yes'
    prop_type_groups = []
    
    if group_prop_types:
        print("Enter groups of PROPERTY TYPEs (e.g., 'Group 1: Residential, Condo; Group 2: Commercial'). Use semicolons to separate groups.")
        groups_input = input("Enter PROPERTY TYPE groups (e.g., Group 1: Residential, Condo; Group 2: Commercial): ").strip()
        if not groups_input:
            raise ValueError("Client 4: No PROPERTY TYPE groups provided")
        
        # Parse groups
        group_list = [group.strip() for group in groups_input.split(';') if group.strip()]
        for i, group in enumerate(group_list, 1):
            prop_types = [pt.strip() for pt in group.split(':')[-1].split(',') if pt.strip()]
            if not prop_types:
                raise ValueError(f"Client 4: Group {i} has no valid PROPERTY TYPEs")
            invalid_prop_types = [pt for pt in prop_types if pt not in valid_prop_types]
            if invalid_prop_types:
                raise ValueError(f"Client 4: Invalid PROPERTY TYPE values in group {i}: {', '.join(invalid_prop_types)}")
            prop_type_groups.append((f"group_{i}", prop_types))
        
        # Check for duplicate PROPERTY TYPEs across groups
        all_prop_types = [pt for _, prop_types in prop_type_groups for pt in prop_types]
        if len(all_prop_types) != len(set(all_prop_types)):
            raise ValueError("Client 4: Duplicate PROPERTY TYPE values found across groups")
    else:
        prop_types_input = input("Enter PROPERTY TYPE values for Client 4 (comma-separated, e.g., Residential, Commercial): ").strip()
        prop_types = [pt.strip() for pt in prop_types_input.split(',') if pt.strip()]
        if not prop_types:
            raise ValueError("Client 4: No valid PROPERTY TYPE values provided")
        invalid_prop_types = [pt for pt in prop_types if pt not in valid_prop_types]
        if invalid_prop_types:
            raise ValueError(f"Client 4: Invalid PROPERTY TYPE values: {', '.join(invalid_prop_types)}")
        prop_type_groups = [(pt, [pt]) for pt in prop_types]
    
    return prop_type_groups

# Function to prompt for Client 5 counties or grouped counties
def prompt_client_5_params(df):
    valid_counties = sorted(df['COUNTY'].dropna().unique())
    if not valid_counties:
        raise ValueError("Client 5: No valid COUNTY values found in data")
    print(f"Available COUNTY values: {', '.join(valid_counties)}")
    
    # Prompt for grouping option
    group_counties = input("Do you want to split Client 5 by grouped counties? (yes/no): ").strip().lower() == 'yes'
    county_groups = []
    
    if group_counties:
        print("Enter groups of counties (e.g., 'Group 1: Wayne, Macomb; Group 2: Oakland'). Use semicolons to separate groups.")
        groups_input = input("Enter county groups (e.g., Group 1: Wayne, Macomb; Group 2: Oakland): ").strip()
        if not groups_input:
            raise ValueError("Client 5: No county groups provided")
        
        # Parse groups
        group_list = [group.strip() for group in groups_input.split(';') if group.strip()]
        for i, group in enumerate(group_list, 1):
            counties = [county.strip() for county in group.split(':')[-1].split(',') if county.strip()]
            if not counties:
                raise ValueError(f"Client 5: Group {i} has no valid counties")
            invalid_counties = [c for c in counties if c not in valid_counties]
            if invalid_counties:
                raise ValueError(f"Client 5: Invalid COUNTY values in group {i}: {', '.join(invalid_counties)}")
            county_groups.append((f"group_{i}", counties))
        
        # Check for duplicate counties across groups
        all_counties = [county for _, counties in county_groups for county in counties]
        if len(all_counties) != len(set(all_counties)):
            raise ValueError("Client 5: Duplicate counties found across groups")
    else:
        counties_input = input("Enter counties for Client 5 (comma-separated, e.g., Wayne, Macomb): ").strip()
        counties = [county.strip() for county in counties_input.split(',') if county.strip()]
        if not counties:
            raise ValueError("Client 5: No valid counties provided")
        invalid_counties = [c for c in counties if c not in valid_counties]
        if invalid_counties:
            raise ValueError(f"Client 5: Invalid COUNTY values: {', '.join(invalid_counties)}")
        county_groups = [(county, [county]) for county in counties]
    
    return county_groups

# Function to prompt for Client 1 split proportions
def prompt_client_1_proportions():
    use_default = input("Do you want to continue with a 50/50 split for Client 1? (yes/no): ").strip().lower() == 'yes'
    if use_default:
        return 50, 50
    
    while True:
        try:
            proportions_input = input("Enter proportions for Template A and Template B (e.g., 60, 40): ").strip()
            prop_a, prop_b = map(float, proportions_input.split(','))
            if prop_a <= 0 or prop_b <= 0:
                raise ValueError("Proportions must be positive")
            if abs(prop_a + prop_b - 100) > 0.001:
                raise ValueError("Proportions must sum to 100")
            return prop_a, prop_b
        except ValueError as e:
            if "must sum to 100" in str(e) or "must be positive" in str(e):
                print(f"Error: {str(e)}")
            else:
                print("Please enter two numbers separated by a comma (e.g., 60, 40)")

# Function for Client 1: Split into two templates with balanced plans and user-specified proportions
def split_client_1(df):
    validate_columns(df, ['FOLIO', 'ACTION PLANS', 'PROPERTY TYPE'], '1')
    valid_plans = ['30 DAYS', '60 DAYS','60 DAYS B', '90 DAYS','90 DAYS B','90 DAYS C']
    if not df['ACTION PLANS'].isin(valid_plans).all():
        invalid_plans = df[~df['ACTION PLANS'].isin(valid_plans)]['ACTION PLANS'].unique()
        raise ValueError(f"Client 1: Invalid ACTION PLANS values: {', '.join(map(str, invalid_plans))}")
    if len(df) < 1000:
        raise ValueError("Client 1: Insufficient records, need at least 1000")
    
    df = df.copy()
    output_dfs = {'client_1_template_a': pd.DataFrame(), 'client_1_template_b': pd.DataFrame()}
    
    # Prompt for proportions
    prop_a, prop_b = prompt_client_1_proportions()
    
    # Calculate plan counts dynamically
    plan_counts = df['ACTION PLANS'].value_counts().to_dict()
    total_records = len(df)
    target_a = int(total_records * prop_a / 100)
    target_b = total_records - target_a  # Ensure exact total match
    
    # Calculate initial expected counts per plan based on proportions
    expected_counts_a = {plan: int(count * prop_a / 100) for plan, count in plan_counts.items()}
    expected_counts_b = {plan: count - expected_counts_a.get(plan, 0) for plan, count in plan_counts.items()}
    
    # Adjust counts to match target_a and target_b exactly
    total_a = sum(expected_counts_a.values())
    total_b = sum(expected_counts_b.values())
    
    # Distribute any shortfall or excess in template_a
    diff_a = target_a - total_a
    if diff_a != 0:
        sorted_plans = sorted(plan_counts.items(), key=lambda x: x[1], reverse=True)  # Sort by count descending
        for plan, count in sorted_plans:
            if diff_a > 0 and expected_counts_a[plan] < count:  # Add records
                addable = min(diff_a, count - expected_counts_a[plan])
                expected_counts_a[plan] += addable
                expected_counts_b[plan] -= addable
                diff_a -= addable
            elif diff_a < 0 and expected_counts_a[plan] > 0:  # Remove records
                removable = min(-diff_a, expected_counts_a[plan])
                expected_counts_a[plan] -= removable
                expected_counts_b[plan] += removable
                diff_a += removable
            if diff_a == 0:
                break
    
    # Verify adjusted counts
    if sum(expected_counts_a.values()) != target_a or sum(expected_counts_b.values()) != target_b:
        raise ValueError(f"Client 1: Adjusted counts do not match targets (A: {sum(expected_counts_a.values())}/{target_a}, B: {sum(expected_counts_b.values())}/{target_b})")
    
    # Split each plan
    for plan in valid_plans:
        plan_df = df[df['ACTION PLANS'] == plan]
        count_a = expected_counts_a.get(plan, 0)
        count_b = expected_counts_b.get(plan, 0)
        if count_a == 0 or count_b == 0:
            warnings.warn(f"Client 1: No contacts for {plan} plan in one template")
        if count_a > len(plan_df) or count_b > len(plan_df):
            raise ValueError(f"Client 1: Insufficient {plan} records for requested proportions")
        
        sampled_a = plan_df.sample(n=count_a, random_state=42) if count_a > 0 else pd.DataFrame()
        plan_df = plan_df.drop(sampled_a.index)
        sampled_b = plan_df.head(count_b) if count_b > 0 else pd.DataFrame()
        
        output_dfs['client_1_template_a'] = pd.concat([output_dfs['client_1_template_a'], sampled_a])
        output_dfs['client_1_template_b'] = pd.concat([output_dfs['client_1_template_b'], sampled_b])
    
    # Verify counts and print plan distribution
    for template, template_df in output_dfs.items():
        actual_counts = template_df['ACTION PLANS'].value_counts().to_dict()
        print(f"Client 1: Plan counts in {template}: {actual_counts}")
        expected_size = target_a if template == 'client_1_template_a' else target_b
        if len(template_df) != expected_size:
            raise ValueError(f"Client 1: {template} has {len(template_df)} records, expected {expected_size}")
    
    # Verify all records used
    total_output = len(output_dfs['client_1_template_a']) + len(output_dfs['client_1_template_b'])
    if total_output != len(df):
        raise ValueError(f"Client 1: Total output records ({total_output}) do not match input ({len(df)})")
    
    # Sort by index to preserve original order
    output_dfs['client_1_template_a'] = output_dfs['client_1_template_a'].sort_index()
    output_dfs['client_1_template_b'] = output_dfs['client_1_template_b'].sort_index()
    
    return output_dfs

# Function for Client 2: Split by odd/even row positions
def split_client_2(df):
    validate_columns(df, ['FOLIO', 'PROPERTY TYPE'], '2')
    if len(df) < 1000:
        raise ValueError("Client 2: Insufficient records, need at least 1000")
    df = df.copy()
    df['Group'] = df.index.map(lambda x: 'Even' if x % 2 == 0 else 'Odd')
    even = df[df['Group'] == 'Even']
    odd = df[df['Group'] == 'Odd']
    even = even.sort_index()
    odd = odd.sort_index()
    return {'client_2_even': even, 'client_2_odd': odd}

# Function for Client 3: Include top X properties in both templates, split remaining into two templates
def split_client_3(df, top_x):
    validate_columns(df, ['FOLIO', 'ACTION PLANS', 'PROPERTY TYPE', 'SCORE'], '3')
    valid_plans = ['30 DAYS', '60 DAYS','60 DAYS B', '90 DAYS','90 DAYS B','90 DAYS C']
    if not df['ACTION PLANS'].isin(valid_plans).all():
        invalid_plans = df[~df['ACTION PLANS'].isin(valid_plans)]['ACTION PLANS'].unique()
        raise ValueError(f"Client 3: Invalid ACTION PLANS values: {', '.join(map(str, invalid_plans))}")
    if not np.issubdtype(df['SCORE'].dtype, np.number):
        raise ValueError("Client 3: SCORE column must be numeric")
    if len(df) < 1000:
        raise ValueError("Client 3: Insufficient records, need at least 1000")
    if top_x > len(df):
        raise ValueError(f"Client 3: Requested top {top_x} properties exceed total records {len(df)}")
    
    df = df.copy()
    output_dfs = {'client_3_template_a': pd.DataFrame(), 'client_3_template_b': pd.DataFrame()}
    
    # Step 1: Select top X properties by SCORE
    top_indices = df['SCORE'].nlargest(top_x).index
    top_df = df.loc[top_indices]
    
    # Initialize templates with top X
    output_dfs['client_3_template_a'] = top_df
    output_dfs['client_3_template_b'] = top_df
    
    # Split remaining records
    remaining_df = df.drop(top_indices)
    remaining_plan_counts = remaining_df['ACTION PLANS'].value_counts().to_dict()
    remaining_total = len(df) - top_x
    target_remaining = remaining_total // 2
    expected_remaining_counts = {plan: count // 2 for plan, count in remaining_plan_counts.items()}
    
    # Handle odd counts
    remainder = remaining_total % 2
    last_plan = valid_plans[-1]
    if remainder and last_plan in expected_remaining_counts:
        expected_remaining_counts[last_plan] += 1
    
    # Split remaining records
    for plan in valid_plans:
        plan_df = remaining_df[remaining_df['ACTION PLANS'] == plan]
        needed_a = expected_remaining_counts.get(plan, 0)
        needed_b = remaining_plan_counts.get(plan, 0) - needed_a
        if needed_a == 0 or needed_b == 0:
            warnings.warn(f"Client 3: No remaining {plan} records for one template")
        
        if needed_a > 0:
            sampled_a = plan_df.sample(n=needed_a, random_state=42) if needed_a <= len(plan_df) else plan_df
            output_dfs['client_3_template_a'] = pd.concat([output_dfs['client_3_template_a'], sampled_a])
            plan_df = plan_df.drop(sampled_a.index)
        if needed_b > 0:
            sampled_b = plan_df.head(needed_b) if needed_b <= len(plan_df) else plan_df
            output_dfs['client_3_template_b'] = pd.concat([output_dfs['client_3_template_b'], sampled_b])
    
    # Verify counts
    target_total = top_x + target_remaining
    for template, template_df in output_dfs.items():
        actual_counts = template_df['ACTION PLANS'].value_counts().to_dict()
        print(f"Client 3: Plan counts in {template}: {actual_counts}")
        if len(template_df) != target_total:
            raise ValueError(f"Client 3: {template} has {len(template_df)} records, expected {target_total}")
    
    # Verify all records used
    total_output = len(output_dfs['client_3_template_a']) + len(output_dfs['client_3_template_b']) - top_x
    if total_output != len(df):
        warnings.warn(f"Client 3: Total output records ({total_output}) do not match input ({len(df)})")
    
    output_dfs['client_3_template_a'] = output_dfs['client_3_template_a'].sort_index()
    output_dfs['client_3_template_b'] = output_dfs['client_3_template_b'].sort_index()
    
    return output_dfs

# Function for Client 4: Split by user-specified PROPERTY TYPE values or groups
def split_client_4(df, prop_type_groups):
    validate_columns(df, ['FOLIO', 'PROPERTY TYPE'], '4')
    valid_prop_types = sorted(df['PROPERTY TYPE'].dropna().unique())
    all_prop_types = [pt for _, prop_types in prop_type_groups for pt in prop_types]
    invalid_types = sorted(set(all_prop_types) - set(valid_prop_types))
    if invalid_types:
        raise ValueError(f"Client 4: Invalid PROPERTY TYPE values in data: {', '.join(invalid_types)}. Valid values are: {', '.join(valid_prop_types)}")
    if len(df) < 1000:
        raise ValueError("Client 4: Insufficient records, need at least 1000")
    df = df.copy()
    output_dfs = {}
    
    for group_name, prop_types in prop_type_groups:
        # Filter records for the PROPERTY TYPEs in this group
        group_df = df[df['PROPERTY TYPE'].isin(prop_types)]
        if len(group_df) == 0:
            warnings.warn(f"Client 4: No records found for group {group_name} with PROPERTY TYPEs {', '.join(prop_types)}")
            continue
        if len(group_df) < 10:
            warnings.warn(f"Client 4: Group {group_name} has {len(group_df)} records, skipping (minimum 10 required)")
            continue
        
        # Validate that all records in group_df have valid PROPERTY TYPEs
        invalid_group_types = sorted(set(group_df['PROPERTY TYPE'].unique()) - set(prop_types))
        if invalid_group_types:
            raise ValueError(f"Client 4: Invalid PROPERTY TYPE values in group {group_name}: {', '.join(invalid_group_types)}")
        
        # Create a single template for this group
        template_name = f'client_4_{group_name.lower().replace(" ", "_")}'
        output_dfs[template_name] = group_df.sort_index()
    
    if not output_dfs:
        raise ValueError("Client 4: No valid groups with sufficient records")
    
    # Verify total records
    total_output = sum(len(df) for df in output_dfs.values())
    if total_output < len(df):
        warnings.warn(f"Client 4: Total output records ({total_output}) less than input ({len(df)}) due to unassigned PROPERTY TYPEs or insufficient records")
    
    return output_dfs

# Function for Client 5: Split by user-specified counties or grouped counties
def split_client_5(df, county_groups):
    validate_columns(df, ['FOLIO', 'COUNTY', 'PROPERTY TYPE'], '5')
    df = df.copy()
    output_dfs = {}
    
    for group_name, counties in county_groups:
        # Filter records for the counties in this group
        group_df = df[df['COUNTY'].isin(counties)]
        if len(group_df) == 0:
            warnings.warn(f"Client 5: No records found for group {group_name} with counties {', '.join(counties)}")
            continue
        if len(group_df) < 10:
            warnings.warn(f"Client 5: Group {group_name} has {len(group_df)} records, skipping (minimum 10 required)")
            continue
        
        # Validate that all records in group_df have valid counties
        invalid_counties = sorted(set(group_df['COUNTY'].unique()) - set(counties))
        if invalid_counties:
            raise ValueError(f"Client 5: Invalid COUNTY values in group {group_name}: {', '.join(invalid_counties)}")
        
        # Create a single template for this group
        template_name = f'client_5_{group_name.lower().replace(" ", "_")}'
        output_dfs[template_name] = group_df.sort_index()
    
    if not output_dfs:
        raise ValueError("Client 5: No valid groups with sufficient records")
    
    # Verify total records
    total_output = sum(len(df) for df in output_dfs.values())
    if total_output < len(df):
        warnings.warn(f"Client 5: Total output records ({total_output}) less than input ({len(df)}) due to unassigned counties or insufficient records")
    
    return output_dfs

# NEW FUNCTION: For Client 6: Split by ACTION PLANS content
def split_client_6(df):
    """Splits DataFrame into separate sheets based on action plan keywords."""
    validate_columns(df, ['FOLIO', 'ACTION PLANS', 'PROPERTY TYPE'], '6')
    df = df.copy()
    output_dfs = {}

    # Define the plan keywords to search for
    plan_keywords = ['30 DAYS', '60 DAYS', '90 DAYS']
    
    # Create a copy of the dataframe to track used records
    remaining_df = df.copy()

    for keyword in plan_keywords:
        # Filter DataFrame for rows where 'ACTION PLANS' contains the keyword
        # Using str.contains() and na=False to handle non-string/NaN values
        # Select from remaining_df to avoid duplicating records
        plan_df = remaining_df[remaining_df['ACTION PLANS'].str.contains(keyword, na=False)]
        
        if not plan_df.empty:
            # Remove these records from remaining_df
            remaining_df = remaining_df.drop(plan_df.index)
            
            # Sanitize sheet name
            sheet_name_key = keyword.replace(' ', '_').lower()
            template_name = f'client_6_{sheet_name_key}'
            output_dfs[template_name] = plan_df.sort_index()
            print(f"Client 6: Found {len(plan_df)} records for ACTION PLANS containing '{keyword}'")
        else:
            warnings.warn(f"Client 6: No records found for ACTION PLANS containing '{keyword}'")

    if not output_dfs:
        raise ValueError("Client 6: No records found for any specified ACTION PLANS keywords (30 DAYS, 60 DAYS, 90 DAYS)")
    
    # Report on any records that didn't match
    if not remaining_df.empty:
        warnings.warn(f"Client 6: {len(remaining_df)} records did not match any ACTION PLANS keyword and were excluded.")

    return output_dfs


# Main function to process multiple Excel files
def process_files():
    # Use hardcoded folder paths
    input_dir = Path(INPUT_FOLDER)
    output_dir = Path(OUTPUT_FOLDER)
    
    # Validate input folder
    if not input_dir.exists() or not input_dir.is_dir():
        raise ValueError(f"Input folder '{INPUT_FOLDER}' does not exist or is not a directory")
    
    # Create output folder if it doesn't exist
    output_dir.mkdir(exist_ok=True)
    
    # Get list of Excel files
    input_files = [f for f in input_dir.glob("*.xlsx") if f.is_file()]
    if not input_files:
        raise ValueError(f"No Excel files found in input folder '{INPUT_FOLDER}'")
    
    # Print process summary for all clients before selection
    print_process_summary()
    
    # MODIFIED: Added client 6 to prompt and validation
    # Prompt user for client selection
    print("Select client to process (1, 2, 3, 4, 5, 6, or 'all' for all clients):")
    client = input("Client (1, 2, 3, 4, 5, 6, or all): ").strip().lower()
    if client not in ['1', '2', '3', '4', '5', '6', 'all']:
        raise ValueError("Invalid client selection. Choose 1, 2, 3, 4, 5, 6, or 'all'.")
    
    all_sheets = []
    
    # Process each Excel file
    for input_file in input_files:
        print(f"\nProcessing file: {input_file}")
        
        # Read input Excel
        try:
            df = pd.read_excel(input_file, engine='openpyxl')
        except Exception as e:
            print(f"Failed to read Excel file {input_file}: {str(e)}")
            continue
        
        # Validate FOLIO and PROPERTY TYPE
        try:
            validate_columns(df, ['FOLIO', 'PROPERTY TYPE'], 'All')
        except ValueError as e:
            print(f"Validation error for {input_file}: {str(e)}")
            continue
        
        # Prompt for Client 3, 4, and 5 parameters if needed
        client_3_top_x = 0
        client_4_prop_type_groups = []
        client_5_county_groups = []
        if client in ['3', 'all']:
            try:
                client_3_top_x = prompt_client_3_params(df)
            except ValueError as e:
                print(f"Client 3 parameter error for {input_file}: {str(e)}")
                continue
        if client in ['4', 'all']:
            try:
                client_4_prop_type_groups = prompt_client_4_params(df)
            except ValueError as e:
                print(f"Client 4 parameter error for {input_file}: {str(e)}")
                continue
        if client in ['5', 'all']:
            try:
                client_5_county_groups = prompt_client_5_params(df)
            except ValueError as e:
                print(f"Client 5 parameter error for {input_file}: {str(e)}")
                continue
        
        # Process splits
        output_dfs = {}
        output_file_name = f"output_{input_file.stem}_client_{client}.xlsx" if client != 'all' else f"output_{input_file.stem}.xlsx"
        output_file = output_dir / output_file_name
        
        try:
            # MODIFIED: Added elif for client 6 and added split_client_6 to the 'all' option
            if client == '1':
                output_dfs.update(split_client_1(df))
            elif client == '2':
                output_dfs.update(split_client_2(df))
            elif client == '3':
                output_dfs.update(split_client_3(df, client_3_top_x))
            elif client == '4':
                output_dfs.update(split_client_4(df, client_4_prop_type_groups))
            elif client == '5':
                output_dfs.update(split_client_5(df, client_5_county_groups))
            elif client == '6':
                output_dfs.update(split_client_6(df))
            else:  # all
                output_dfs.update(split_client_1(df))
                output_dfs.update(split_client_2(df))
                output_dfs.update(split_client_3(df, client_3_top_x))
                output_dfs.update(split_client_4(df, client_4_prop_type_groups))
                output_dfs.update(split_client_5(df, client_5_county_groups))
                output_dfs.update(split_client_6(df))
            
            # Write to Excel
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                for sheet_name, sheet_df in output_dfs.items():
                    sheet_name = sheet_name[:31]  # Excel sheet names <= 31 chars
                    sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print(f"Generated Excel file with {len(output_dfs)} sheets: {output_file}")
            
            # Print final record counts and plan distributions for Client 1 and Client 3
            if client in ['1', 'all'] and 'client_1_template_a' in output_dfs:
                print(f"Client 1: Template A has {len(output_dfs['client_1_template_a'])} records")
                print(f"Client 1: Template B has {len(output_dfs['client_1_template_b'])} records")
                print(f"Client 1: Plan counts in client_1_template_a: {output_dfs['client_1_template_a']['ACTION PLANS'].value_counts().to_dict()}")
                print(f"Client 1: Plan counts in client_1_template_b: {output_dfs['client_1_template_b']['ACTION PLANS'].value_counts().to_dict()}")
            if client in ['3', 'all']:
                for template, template_df in output_dfs.items():
                    if template.startswith('client_3'):
                        print(f"Client 3: {template} has {len(template_df)} records")
                        print(f"Client 3: Plan counts in {template}: {template_df['ACTION PLANS'].value_counts().to_dict()}")
            
            all_sheets.append((input_file.name, list(output_dfs.keys())))
        
        except ValueError as e:
            print(f"Error processing {input_file}: {str(e)}")
            continue
        except Exception as e:
            print(f"Unexpected error processing {input_file}: {str(e)}")
            continue
    
    if not all_sheets:
        print("No files were successfully processed")
    else:
        print("\nSummary of processed files:")
        for file_name, sheets in all_sheets:
            print(f"File: {file_name}, Generated sheets: {sheets}")
    
    return all_sheets

# Example usage
if __name__ == "__main__":
    try:
        sheets = process_files()
        if sheets:
            print("\nAll files processed successfully")
    except ValueError as e:
        print(f"Error: {str(e)}")
    except Exception as e:
        print(f"Unexpected error: {str(e)}")