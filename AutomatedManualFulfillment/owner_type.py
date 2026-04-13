import pandas as pd

# --- Define Keyword Lists Based on Spark owner_name_conditions ---
# Order matches owner_name_conditions to preserve Spark hierarchy

# Rule: CheckWithSpaces (matches % keyword % patterns)
rules_check_with_spaces = [
    ("co tr", "Trust"),
    ("united states of america", "Non Sellers"),
    ("postal service", "Non Sellers"),
    ("llc", "Company"),
    ("l l c", "Company"),
    ("inc", "Company"),
    ("company", "Company"),
    ("corp", "Company"),
    ("enterprise", "Company"),
    ("partnership", "Company"),
    ("group", "Company"),
    ("corporation", "Company"),
    ("invest", "Company"),
    ("holdings", "Company"),
    ("ltd", "Company"),
    ("limited", "Company"),
    ("bank", "Non Sellers"),
    ("center", "Non Sellers"),
    ("school", "Non Sellers"),
    ("university", "Non Sellers"),
    ("church", "Company"),  # Mapped to Company (no Religious Organization)
    ("churc", "Company"),
    ("baptist", "Company"),
    ("episcopal", "Company"),
    ("apostol", "Company"),
    ("minister", "Company"),
    ("iglesia", "Company"),
    ("orthodox", "Company"),
    ("missionary", "Company"),
    ("chrch", "Company"),
    ("diocese", "Company"),
    ("bishop of", "Company"),
    ("ministry", "Company"),
    ("ministries", "Company"),
    ("fellowship", "Company"),
    ("believer", "Company"),
    ("christian center", "Company"),
    ("worship", "Company"),
    ("pentecostal", "Company"),
    ("islamic", "Company"),
    ("methodist", "Non Sellers"),
    ("trstee", "Trust"),
    ("rev tr", "Trust"),
    ("rev liv tr", "Trust"),
    ("trste", "Trust"),
    ("trust", "Trust"),
    ("trs", "Non Sellers"),
    ("dist", "Non Sellers"),
    ("hoa", "Company"),
    ("condo", "Company"),
    ("home", "Company"),
    ("service", "Company"),
    ("dev", "Company"),
    ("plaza", "Company"),
    ("bank of", "Non Sellers"),
    ("td bank", "Non Sellers"),
    ("us bank", "Non Sellers"),
    ("u s bank", "Non Sellers"),
    ("management", "Company"),
    ("national", "Non Sellers"),
    ("mortgage", "Non Sellers"),
    ("mgmt", "Non Sellers"),
    ("loans", "Non Sellers"),
    ("finance", "Non Sellers"),
    ("univers", "Non Sellers"),
    ("united", "Non Sellers"),
    ("authority", "Non Sellers"),
    ("congr", "Non Sellers"),
    ("florida", "Non Sellers"),
    ("college", "Non Sellers"),
    ("county", "Non Sellers"),
    ("state of", "Non Sellers"),
    ("city", "Non Sellers"),
    ("broward", "Non Sellers"),
    ("miami-dade", "Non Sellers"),
    ("dept", "Non Sellers"),
    ("natl", "Non Sellers"),
    ("u s a", "Non Sellers"),
    ("everglades", "Non Sellers"),
    ("public", "Non Sellers"),
    ("congregation", "Non Sellers"),
    ("coalition", "Non Sellers"),
    ("outreach", "Non Sellers"),
    ("assemble", "Non Sellers"),
    ("alliance", "Non Sellers"),
    ("assembly", "Non Sellers"),
    ("independent", "Non Sellers"),
    ("traditional", "Non Sellers"),
    ("secretary", "Non Sellers"),
    ("of the", "Non Sellers"),
    ("humanity", "Non Sellers"),
    ("federal", "Non Sellers"),
    ("federation", "Non Sellers"),
    ("drainage", "Non Sellers"),
    ("energy", "Non Sellers"),
    ("railroad", "Non Sellers"),
    ("cemeter", "Non Sellers"),
    ("develop", "Company"),
    ("assoc", "Company"),
    ("companies", "Company"),
    ("inversiones", "Company"),
    ("community", "Company"),
    ("construction", "Company"),
    ("agency", "Company"),
    ("professional", "Company"),
    ("properties", "Company"),
    ("housing", "Company"),
    ("international", "Company"),
    ("supplies", "Company"),
    ("equip", "Company"),
    ("realty", "Company"),
    ("rental", "Company"),
    ("product", "Company"),
    ("wholesale", "Company"),
    ("apartment", "Company"),
    ("consult", "Company"),
    ("real estate", "Company"),
    ("shopping", "Company"),
    ("trading", "Company"),
    ("condos", "Company"),
    ("apts", "Company"),
    ("village", "Company"),
    ("family", "Company"),
    ("business", "Company"),
    ("solutions", "Company"),
    ("detailing", "Company"),
    ("fashion", "Company"),
    ("commercial", "Company"),
    ("studio", "Company"),
    ("design", "Company"),
    ("barbershop", "Company"),
    ("rescue", "Company"),
    ("adoption", "Company"),
    ("academy", "Company"),
    ("repair", "Company"),
    ("custom", "Company"),
    ("industries", "Company"),
    ("cemetries", "Company"),
    ("lllp", "Company"),
    ("1", "Company"), ("2", "Company"), ("3", "Company"),
    ("4", "Company"), ("5", "Company"), ("6", "Company"),
    ("7", "Company"), ("8", "Company"), ("9", "Company"), ("0", "Company"),
    ("prtnrsp", "Company"), ("ptnrhp", "Company"), ("ptnrshp", "Company"),
    ("prtnrshp", "Company"), ("ptnshp", "Company"),
    ("operations", "Company"),
    ("realtor", "Company"),
    ("venture", "Company"),
    ("resident", "Company"),
    ("warehouse", "Company"),
    ("storage", "Company"),
    ("communication", "Company"),
    ("healthcare", "Company"),
    ("medical", "Company"),
    ("citiside", "Company"),
    ("condominium", "Company"),
    ("builders", "Company"),
    ("communities", "Company"),
    ("homeowners", "Company"),
    ("timeshare", "Company"),
]

# Rule: CheckContains (matches %keyword% patterns)
rules_check_contains = [
    ("office", "Unknown"),
    ("available from", "Unknown"),
    ("churchill", "Individual"),
    ("upchurch", "Individual"),
]

# Rule: CheckEndsWith (matches %keyword patterns)
rules_check_ends_with = [
    ("usa", "Non Sellers"),
    ("tr", "Trust"),
    ("pa", "Company"),
    ("sa", "Company"),
    ("s a", "Company"),
    ("l c", "Company"),
    ("lp", "Company"),
    ("en", "Company"),
    ("co", "Company"),
    ("gp", "Company"),
]

# Rule: CheckStartsWith (matches keyword% patterns)
rules_check_starts_with = [
    ("co tr", "Trust"),
    ("llc", "Company"),
    ("l l c", "Company"),
    ("inc", "Company"),
    ("corp", "Company"),
    ("bank", "Non Sellers"),
    ("church", "Company"),
    ("churc", "Company"),
    ("baptist", "Company"),
    ("episcopal", "Company"),
    ("apostol", "Company"),
    ("minister", "Company"),
    ("iglesia", "Company"),
    ("missionary", "Company"),
    ("bishop of", "Company"),
    ("dist", "Company"),
    ("hoa", "Company"),
    ("condo", "Company"),
    ("home", "Company"),
    ("dev", "Company"),
    ("plaza", "Company"),
]

# --- Helper Functions for Rules ---

def apply_check_with_spaces(name_full, rules):
    """Matches patterns like % keyword %"""
    if not isinstance(name_full, str) or not name_full.strip():
        return None
    name_lower = name_full.lower()
    for keyword, owner_type in rules:
        k_lower = keyword.lower()
        if f" {k_lower} " in name_lower or \
           name_lower == k_lower or \
           name_lower.startswith(f"{k_lower} ") or \
           name_lower.endswith(f" {k_lower}"):
            return owner_type
    return None

def apply_check_contains(name_full, rules):
    """Matches patterns like %keyword%"""
    if not isinstance(name_full, str) or not name_full.strip():
        return None
    name_lower = name_full.lower()
    for keyword, owner_type in rules:
        k_lower = keyword.lower()
        if k_lower in name_lower:
            return owner_type
    return None

def apply_check_ends_with(name_full, rules):
    """Matches patterns like %keyword"""
    if not isinstance(name_full, str) or not name_full.strip():
        return None
    name_lower_stripped = name_full.lower().strip()
    for keyword, owner_type in rules:
        k_lower = keyword.lower()
        if name_lower_stripped.endswith(k_lower):
            return owner_type
    return None

def apply_check_starts_with(name_full, rules):
    """Matches patterns like keyword%"""
    if not isinstance(name_full, str) or not name_full.strip():
        return None
    name_lower = name_full.lower()
    for keyword, owner_type in rules:
        k_lower = keyword.lower()
        if name_lower.startswith(k_lower):
            return owner_type
    return None

def apply_check_by_address_adapted(first_name, last_name, property_type_excel_val):
    """Adapted from PDF[cite:16]"""
    first_name_str = str(first_name) if pd.notna(first_name) else ""
    last_name_str = str(last_name) if pd.notna(last_name) else ""
    property_type_str = str(property_type_excel_val).upper().strip() if pd.notna(property_type_excel_val) else ""
    is_first_name_present = first_name_str.strip() != ""
    is_last_name_present = last_name_str.strip() != ""
    valid_use_types_for_individual = ["SFH", "TOWNHOUSE", "LAND", "CONDO"]
    if is_first_name_present and is_last_name_present and property_type_str in valid_use_types_for_individual:
        return "INDIVIDUAL"
    return None

# --- Main Processing Function ---
def process_owner_data(input_excel_path, output_excel_path):
    """Processes Excel file using Spark's name-based rules"""
    try:
        df = pd.read_excel(input_excel_path, engine='openpyxl')
        print(f"Successfully read {len(df)} rows from {input_excel_path}")
    except FileNotFoundError:
        print(f"Error: Input file not found at {input_excel_path}")
        return
    except Exception as e:
        print(f"Error reading Excel file '{input_excel_path}': {e}")
        return

    df_output = df.copy()
    df_original_columns = df.columns.tolist()
    owner_types_results = []

    required_columns = ["OWNER FULL NAME", "OWNER FIRST NAME", "OWNER LAST NAME", "PROPERTY TYPE"]
    missing_columns = [col for col in required_columns if col not in df_original_columns]
    if missing_columns:
        print(f"Error: Input file is missing required columns: {missing_columns}")
        return

    for index, row in df_output.iterrows():
        current_owner_type = None

        # Name-based rules (no owner_type_dict)
        owner_full_name = str(row.get("OWNER FULL NAME", "")) if pd.notna(row.get("OWNER FULL NAME")) else ""
        owner_first_name = row.get("OWNER FIRST NAME")
        owner_last_name = row.get("OWNER LAST NAME")
        property_type = row.get("PROPERTY TYPE")

        if current_owner_type is None:
            current_owner_type = apply_check_with_spaces(owner_full_name, rules_check_with_spaces)
        if current_owner_type is None:
            current_owner_type = apply_check_contains(owner_full_name, rules_check_contains)
        if current_owner_type is None:
            current_owner_type = apply_check_ends_with(owner_full_name, rules_check_ends_with)
        if current_owner_type is None:
            current_owner_type = apply_check_starts_with(owner_full_name, rules_check_starts_with)
        if current_owner_type is None:
            current_owner_type = apply_check_by_address_adapted(owner_first_name, owner_last_name, property_type)
        if current_owner_type is None:
            current_owner_type = "UNKNOWN"

        owner_types_results.append(current_owner_type)

    df_output["owner type"] = owner_types_results
    df_output["property status"] = ""

    final_columns_ordered = [col for col in df_original_columns if col not in ["owner type", "property status"]]
    final_columns_ordered.extend(["owner type", "property status"])
    df_output = df_output[final_columns_ordered]

    try:
        df_output.to_excel(output_excel_path, index=False, engine='openpyxl')
        print(f"Processing complete. Output with {len(df_output)} rows saved to {output_excel_path}")
        print(f"Input file '{input_excel_path}' remains unchanged.")
    except Exception as e:
        print(f"Error writing Excel file '{output_excel_path}': {e}")

# --- Example Usage ---
if __name__ == "__main__":
    actual_input_file = r"C:\Users\LENOVO\Documents\8020\AutomatedManualFulfillment\Rejected_Properties_Output.xlsx"
    actual_output_file = r"C:\Users\LENOVO\Documents\8020\AutomatedManualFulfillment\Processed_Rejected_Properties.xlsx"
    print(f"\n--- Running with Actual Data ---")
    print(f"Input: {actual_input_file} (will not be modified)")
    print(f"Output: {actual_output_file}\n")
    process_owner_data(actual_input_file, actual_output_file)