import csv
import pandas as pd
from datetime import datetime, timedelta
import re
import os
from geopy.geocoders import Nominatim
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
import time
from tqdm import tqdm

# Helper function to convert Excel serial date to datetime
def excel_date_to_datetime(excel_date):
    if isinstance(excel_date, (int, float)):
        base_date = datetime(1899, 12, 30)
        return base_date + timedelta(days=excel_date)
    return excel_date

# Phase 1 Functions
def detect_date_format_phase1(date_input):
    if isinstance(date_input, datetime):
        return date_input.strftime('%m/%d/%Y')  # Changed to mm/dd/yyyy
    if isinstance(date_input, (int, float)):
        try:
            dt = excel_date_to_datetime(date_input)
            return dt.strftime('%m/%d/%Y')
        except:
            return "Invalid Date"
    if isinstance(date_input, str):
        date_formats = ['%m/%d/%y', '%m/%d/%Y', '%Y-%m-%d', '%d/%m/%y', '%d/%m/%Y']
        for fmt in date_formats:
            try:
                return datetime.strptime(date_input, fmt).strftime('%m/%d/%Y')
            except ValueError:
                continue
    return "Invalid Date"

def is_accepted_address_format_phase1(address):
    address = address.strip()
    parts = address.split()
    if len(parts) >= 3:
        last_part = parts[-1]
        zip_match = re.match(r'^\d{5}(-\d{4})?$', last_part)
        if zip_match and len(parts[-2]) == 2 and parts[-2].isalpha():
            return True
        elif len(parts[-1]) == 2 and parts[-1].isalpha():
            return True
    if ',' in address:
        last_comma_index = address.rfind(',')
        if last_comma_index != -1:
            after_comma = address[last_comma_index + 1:].strip()
            after_parts = after_comma.split()
            if len(after_parts) >= 1:
                if len(after_parts) > 1 and re.match(r'^\d{5}(-\d{4})?$', after_parts[-1]) and len(after_parts[-2]) == 2 and after_parts[-2].isalpha():
                    return True
                elif len(after_parts[0]) == 2 and after_parts[0].isalpha():
                    return True
    return False

def normalize_address_phase1(address):
    address = address.strip()
    zip_code = ""
    if ',' in address:
        last_comma_index = address.rfind(',')
        if last_comma_index != -1:
            before_comma, after_comma = address[:last_comma_index].strip(), address[last_comma_index + 1:].strip()
            after_parts = after_comma.split()
            if len(after_parts) >= 2:
                if re.match(r'^\d{5}(-\d{4})?$', after_parts[-1]) and len(after_parts[-2]) == 2 and after_parts[-2].isalpha():
                    zip_code = after_parts[-1]
                    state = after_parts[-2]
                    city = ' '.join(after_parts[:-2]) if len(after_parts) > 2 else ""
                    street = before_comma
                    return street, city, state, zip_code
                elif len(after_parts[0]) == 2 and after_parts[0].isalpha():
                    state = after_parts[0]
                    zip_code = after_parts[1] if len(after_parts) > 1 and re.match(r'^\d{5}(-\d{4})?$', after_parts[1]) else ""
                    city = ' '.join(after_parts[1:]) if len(after_parts) > 1 and not zip_code else ""
                    street = before_comma
                    return street, city, state, zip_code
    parts = address.split()
    if len(parts) >= 3:
        last_part = parts[-1]
        zip_match = re.match(r'^\d{5}(-\d{4})?$', last_part)
        if zip_match and len(parts[-2]) == 2 and parts[-2].isalpha():
            zip_code = last_part
            state = parts[-2]
            city = parts[-3] if len(parts) > 3 else ""
            street = ' '.join(parts[:-3])
            return street, city, state, zip_code
        elif len(parts[-1]) == 2 and parts[-1].isalpha():
            state = parts[-1]
            city = parts[-2] if len(parts) > 2 else ""
            street = ' '.join(parts[:-2])
            return street, city, state, zip_code
    return address.rstrip(', ').rstrip(), "", "", ""

# Phase 2 Functions
def detect_date_format_phase2(date_input):
    if isinstance(date_input, datetime):
        return date_input.strftime('%m/%d/%Y')  # Changed to mm/dd/yyyy
    if isinstance(date_input, (int, float)):
        try:
            dt = excel_date_to_datetime(date_input)
            return dt.strftime('%m/%d/%Y')
        except:
            return "Invalid Date"
    if isinstance(date_input, str):
        date_formats = ['%m/%d/%y', '%m/%d/%Y', '%Y-%m-%d', '%d/%m/%y', '%d/%m/%Y']
        for fmt in date_formats:
            try:
                return datetime.strptime(date_input, fmt).strftime('%m/%d/%Y')
            except ValueError:
                continue
    return "Invalid Date"

def is_accepted_address_format_phase2(address):
    address = address.strip()
    parts = address.split()
    if len(parts) >= 3:
        last_part = parts[-1]
        zip_match = re.match(r'^\d{5}(-\d{4})?$', last_part)
        if zip_match and len(parts[-2]) == 2 and parts[-2].isalpha():
            return True
        elif len(parts[-1]) == 2 and parts[-1].isalpha():
            return True
    if ',' in address:
        last_comma_index = address.rfind(',')
        if last_comma_index != -1:
            after_comma = address[last_comma_index + 1:].strip()
            after_parts = after_comma.split()
            if len(after_parts) >= 1:
                if len(after_parts) > 1 and re.match(r'^\d{5}(-\d{4})?$', after_parts[-1]) and len(after_parts[-2]) == 2 and after_parts[-2].isalpha():
                    return True
                elif len(after_parts[0]) == 2 and after_parts[0].isalpha():
                    return True
    return False

def normalize_address_phase2(address):
    address = address.strip()
    zip_code = ""
    if ',' in address:
        last_comma_index = address.rfind(',')
        if last_comma_index != -1:
            before_comma, after_comma = address[:last_comma_index].strip(), address[last_comma_index + 1:].strip()
            after_parts = after_comma.split()
            if len(after_parts) >= 1:
                if len(after_parts) > 1 and re.match(r'^\d{5}(-\d{4})?$', after_parts[-1]) and len(after_parts[-2]) == 2 and after_parts[-2].isalpha():
                    zip_code = after_parts[-1]
                    state = after_parts[-2]
                    city = ' '.join(after_parts[:-2]) if len(after_parts) > 2 else ""
                    street = before_comma
                    return street, city, state, zip_code
                elif len(after_parts[0]) == 2 and after_parts[0].isalpha():
                    state = after_parts[0]
                    zip_code = after_parts[1] if len(after_parts) > 1 and re.match(r'^\d{5}(-\d{4})?$', after_parts[1]) else ""
                    city = ' '.join(after_parts[1:]) if len(after_parts) > 1 and not zip_code else ""
                    street = before_comma
                    return street, city, state, zip_code
    parts = address.split()
    if len(parts) >= 3:
        last_part = parts[-1]
        zip_match = re.match(r'^\d{5}(-\d{4})?$', last_part)
        if zip_match and len(parts[-2]) == 2 and parts[-2].isalpha():
            zip_code = last_part
            state = parts[-2]
            city = parts[-3] if len(parts) > 3 else ""
            street = ' '.join(parts[:-3])
            return street, city, state, zip_code
        elif len(parts[-1]) == 2 and parts[-1].isalpha():
            state = parts[-1]
            city = parts[-2] if len(parts) > 2 else ""
            street = ' '.join(parts[:-2])
            return street, city, state, zip_code
    return address.rstrip(', ').rstrip(), "", "", ""

# Phase 3 Function
def normalize_address_phase3(address, date="Invalid Date", revenue=0.0):
    address = address.strip()
    parts = address.split()
    street, city, state, zip_code = "", "", "", ""

    if len(parts) >= 3:
        last_part = parts[-1]
        zip_match = re.match(r'^\d{5}(-\d{4})?$', last_part)
        if zip_match and len(parts[-2]) == 2 and parts[-2].isalpha():
            zip_code = last_part
            state = parts[-2]
            city = parts[-3] if len(parts) > 3 else ""
            street = ' '.join(parts[:-3])
        elif len(parts[-1]) == 2 and parts[-1].isalpha():
            state = parts[-1]
            city = parts[-2] if len(parts) > 2 else ""
            street = ' '.join(parts[:-2])
        else:
            street = address
    else:
        street = address

    return [street, city, state, date, revenue, zip_code]

# Phase 4 Function
def normalize_address_phase4(address, date="Invalid Date", revenue=0.0):
    address = address.strip()
    parts = address.split()
    street, city, state, zip_code = "", "", "", ""

    if len(parts) >= 3:
        if re.match(r'^\d{5}(-\d{4})?$', parts[-1]) and len(parts[-2]) == 2 and parts[-2].isalpha():
            zip_code = parts[-1]
            state = parts[-2]
            city = parts[-3] if len(parts) > 3 else ""
            street = ' '.join(parts[:-3])
        elif len(parts[-1]) == 2 and parts[-1].isalpha():
            state = parts[-1]
            city = parts[-2] if len(parts) > 2 else ""
            street = ' '.join(parts[:-2])
        else:
            street = address
    else:
        street = address

    return [street, city, state, date, revenue, zip_code]

# Geocoding Functions
def get_zipcode_nominatim(partial_address):
    geolocator = Nominatim(user_agent="zipcode_finder", timeout=10)
    
    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(multiplier=1, min=1, max=10),
        retry=retry_if_exception_type((ConnectionError, TimeoutError))
    )
    def _geocode_with_retry():
        time.sleep(1)
        location = geolocator.geocode(partial_address, addressdetails=True)
        if location and location.raw.get('address'):
            postcode = location.raw['address'].get('postcode')
            if postcode:
                return postcode.split('-')[0]  # Return only 5-digit ZIP
        return None

    try:
        return _geocode_with_retry()
    except Exception as e:
        return None

def clean_address(address):
    match = re.match(r'^(\d+)[,\s]*(.+)', address)
    if match:
        number, rest = match.groups()
        return f"{number} {rest.strip()}"
    return address

def is_numeric(value):
    if pd.isna(value):
        return False
    try:
        float(str(value).replace('.0', ''))
        return True
    except ValueError:
        return False

def has_zip(value):
    return not pd.isna(value) and str(value).strip() not in ["", "nan"]

def geocode_df(df, output_file):
    start_time = time.time()
    df.columns = [col.lower() for col in df.columns]

    required_columns = ['property address', 'property city', 'property state', 'property zip code']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")

    zips_before = df['property zip code'].apply(has_zip).sum()
    print(f"Total ZIP codes before geocoding: {zips_before}")

    def combine_address(row):
        parts = [str(row['property address']), str(row['property city']), str(row['property state'])]
        valid_parts = [part for part in parts if part and part != 'nan']
        return ', '.join(valid_parts) if valid_parts else "Incomplete address"

    df['full_address'] = df.apply(combine_address, axis=1)
    df['cleaned_address'] = df['full_address'].apply(clean_address)

    print(f"\nGeocoding data...")
    tqdm.pandas()

    mask = ~df['property zip code'].apply(is_numeric) & (df['cleaned_address'].str.contains(',', na=False))
    print(f"Rows needing geocoding: {mask.sum()}")

    if mask.sum() > 0:
        geocoded_zips = df.loc[mask, 'cleaned_address'].progress_apply(get_zipcode_nominatim)
        geocoded_zips = pd.to_numeric(geocoded_zips, errors='coerce')
        df.loc[mask, 'property zip code'] = geocoded_zips

    zips_after = df['property zip code'].apply(has_zip).sum()
    print(f"Total ZIP codes after geocoding: {zips_after}")

    df = df.drop(columns=['full_address', 'cleaned_address'])

    df.to_excel(output_file, index=False)

    end_time = time.time()
    execution_time = end_time - start_time

    print(f"\nResults saved to: '{output_file}'")
    print(f"Execution time: {execution_time:.2f} seconds")
    print(f"Total rows processed: {len(df)}")
    print(f"Rows geocoded: {mask.sum()}")
    print(f"ZIP codes before: {zips_before}, ZIP codes after: {zips_after}")
    if zips_before > zips_after:
        print(f"Warning: Lost {zips_before - zips_after} ZIP codes during geocoding")
    elif zips_after > zips_before:
        print(f"Added {zips_after - zips_before} ZIP codes during geocoding")

    return df

# Main Normalization and Geocoding Function
def normalize_and_geocode_csv(input_file, output_folder):
    data = []
    headers = []
    
    file_extension = os.path.splitext(input_file)[1].lower()
    print(f"Detected file extension: {file_extension}")
    
    if file_extension == '.csv':
        with open(input_file, 'r', encoding='utf-8') as file:
            reader = csv.reader(file)
            headers = next(reader)
            data = [row for row in reader]
    elif file_extension == '.xlsx':
        print(f"Reading Excel file: {input_file}")
        df = pd.read_excel(input_file)
        headers = df.columns.tolist()
        data = [list(row) for row in df.itertuples(index=False)]
    else:
        raise ValueError(f"Unsupported file type: {file_extension}. Only .csv and .xlsx are supported.")

    # Find column indices for Address, Date, and Revenue
    address_col = None
    date_col = None
    revenue_col = None

    for i, header in enumerate(headers):
        header_lower = header.lower()
        if 'address' in header_lower:
            address_col = i
        if 'date' in header_lower:
            date_col = i
        if 'revenue' in header_lower or 'sale' in header_lower or 'price' in header_lower:
            revenue_col = i

    # Fallbacks if columns are not found
    if address_col is None:
        print("Warning: No 'Address' header found, defaulting to first column.")
        address_col = 0
    if date_col is None:
        print("Warning: No 'Date' header found, defaulting to second column.")
        date_col = 1 if len(headers) > 1 else None
    if revenue_col is None:
        print("Warning: No 'Revenue', 'Sale', or 'Price' header found, defaulting to third column.")
        revenue_col = 2 if len(headers) > 2 else None

    # Phase 1
    normalized_data_phase1, unnormalized_data_phase1 = [], []
    normalized_count_phase1, skipped_count_phase1 = 0, 0

    for row in data:
        # Extract address
        address = str(row[address_col]) if address_col is not None and len(row) > address_col else ""

        # Extract and process date
        date = "Invalid Date"
        if date_col is not None and len(row) > date_col:
            date = detect_date_format_phase1(row[date_col]) or "Invalid Date"

        # Extract and process revenue
        revenue = 0.0
        if revenue_col is not None and len(row) > revenue_col:
            revenue_str = str(row[revenue_col]).strip()
            cleaned_revenue = re.sub(r'[^\d.]', '', revenue_str)
            revenue = float(cleaned_revenue) if cleaned_revenue else 0.0

        if is_accepted_address_format_phase1(address):
            street, city, state, zip_code = normalize_address_phase1(address)
            normalized_data_phase1.append([street, city, state, date, revenue, zip_code])
            normalized_count_phase1 += 1
        else:
            unnormalized_data_phase1.append([address, "", "", date, revenue, ""])
            skipped_count_phase1 += 1

    # Phase 2
    normalized_data_phase2, unnormalized_data_phase2 = [], []
    normalized_count_phase2, skipped_count_phase2 = 0, 0

    for row in unnormalized_data_phase1:
        address = row[0]
        if is_accepted_address_format_phase2(address):
            street, city, state, zip_code = normalize_address_phase2(address)
            normalized_data_phase2.append([street, city, state, row[3], row[4], zip_code])
            normalized_count_phase2 += 1
        else:
            unnormalized_data_phase2.append(row)
            skipped_count_phase2 += 1

    # Phase 3
    normalized_data_phase3, unnormalized_data_phase3 = [], []
    normalized_count_phase3, skipped_count_phase3 = 0, 0

    for row in unnormalized_data_phase2:
        address = row[0]
        normalized_row = normalize_address_phase3(address, row[3], row[4])
        if normalized_row[0] != address:
            normalized_data_phase3.append(normalized_row)
            normalized_count_phase3 += 1
        else:
            unnormalized_data_phase3.append(row)
            skipped_count_phase3 += 1

    # Phase 4
    normalized_data_phase4, unnormalized_data_phase4 = [], []
    normalized_count_phase4, skipped_count_phase4 = 0, 0

    for row in unnormalized_data_phase3:
        address = row[0]
        normalized_row = normalize_address_phase4(address, row[3], row[4])
        if normalized_row[0] != address:
            normalized_data_phase4.append(normalized_row)
            normalized_count_phase4 += 1
        else:
            unnormalized_data_phase4.append([row[0], row[1], row[2], row[3], row[4], ""])
            skipped_count_phase4 += 1

    # Combine all data
    all_data = normalized_data_phase1 + normalized_data_phase2 + normalized_data_phase3 + normalized_data_phase4 + unnormalized_data_phase4
    
    # Define output headers
    headers = ['Property Address', 'Property City', 'Property State', 'Date', 'Revenue', 'Property Zip code']
    
    # Create DataFrame
    df = pd.DataFrame(all_data, columns=headers)
    
    # Sort by 'Property Zip code' in descending order, with empty values at the bottom
    df['Property Zip code'] = df['Property Zip code'].replace('', pd.NA)
    df = df.sort_values(by='Property Zip code', ascending=False, na_position='last')

    # Summary (before geocoding)
    print(f"Normalization complete for {os.path.basename(input_file)}!")
    print(f"Phase 1 - Properties normalized: {normalized_count_phase1}")
    print(f"Phase 1 - Properties kept un-normalized: {skipped_count_phase1}")
    print(f"Phase 2 - Properties normalized: {normalized_count_phase2}")
    print(f"Phase 2 - Properties kept un-normalized: {skipped_count_phase2}")
    print(f"Phase 3 - Properties normalized: {normalized_count_phase3}")
    print(f"Phase 3 - Properties kept un-normalized: {skipped_count_phase3}")
    print(f"Phase 4 - Properties normalized: {normalized_count_phase4}")
    print(f"Phase 4 - Properties kept un-normalized: {skipped_count_phase4}")
    print(f"Total properties processed: {len(data)}")

    # Geocode the normalized DataFrame
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    output_file = os.path.join(output_folder, f'normalized_geocoded_{base_name}.xlsx')
    df = geocode_df(df, output_file)

    return df, normalized_count_phase1, normalized_count_phase2, normalized_count_phase3, normalized_count_phase4, skipped_count_phase4

# Main Execution
try:
    input_folder = r"C:\Users\LENOVO\Documents\8020\py_scripts\normalization_code\input file"
    output_folder = r"C:\Users\LENOVO\Documents\8020\py_scripts\normalization_code\output file"

    for filename in os.listdir(input_folder):
        input_file = os.path.join(input_folder, filename)
        if os.path.isfile(input_file):
            print(f"Processing file: {filename}")
            result, norm_count_p1, norm_count_p2, norm_count_p3, norm_count_p4, skipped_count_p4 = normalize_and_geocode_csv(
                input_file, output_folder
            )
            print(f"Finished processing {filename}\n")

except FileNotFoundError:
    print("Error: Input folder not found. Please ensure the directory exists.")
except Exception as e:
    print(f"An error occurred: {str(e)}")