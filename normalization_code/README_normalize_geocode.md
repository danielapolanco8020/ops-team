# Address Normalizer & Geocoder — `normalize_geocode.py`

Processes property records from CSV or Excel files, parses raw address strings into structured components through a 4-phase normalization pipeline, and fills in missing ZIP codes via geocoding using the OpenStreetMap Nominatim API.

---

## What It Does

1. Reads all files from the input folder (`.csv` or `.xlsx`)
2. Auto-detects the address, date, and revenue columns by header name
3. Runs each address through a **4-phase normalization pipeline** to extract street, city, state, and ZIP
4. Standardizes all dates to `MM/DD/YYYY` format (supports Excel serial dates and multiple string formats)
5. Cleans revenue values by stripping currency symbols and non-numeric characters
6. Geocodes rows still missing a ZIP code using the **Nominatim API** (OpenStreetMap)
7. Saves a formatted output Excel file per input file, sorted by ZIP code

---

## Setup

**Requirements:** Python 3.x with the following packages:

```bash
pip install pandas openpyxl geopy tenacity tqdm
```

> **Note:** Geocoding uses the free Nominatim API (OpenStreetMap). It enforces a **1-second delay per request** to comply with usage limits. Large files with many missing ZIPs will take significant time to process.

---

## Configuration

Open `normalize_geocode.py` and update the two folder paths at the bottom:

```python
input_folder  = r"C:\path\to\input file"
output_folder = r"C:\path\to\output file"
```

---

## Usage

```bash
python normalize_geocode.py
```

All files in the input folder are processed automatically. No interactive input is required.

---

## Required Input Columns

The script auto-detects columns by scanning headers for these keywords:

| Keyword Matched | Maps To | Used For |
|---|---|---|
| Contains `address` | Address column | Normalization & geocoding |
| Contains `date` | Date column | Date standardization |
| Contains `revenue`, `sale`, or `price` | Revenue column | Numeric cleaning |

If a column is not found by keyword, the script falls back to positional defaults (1st, 2nd, 3rd columns respectively) and prints a warning.

**Supported input formats:** `.csv`, `.xlsx`

---

## Normalization Pipeline

Raw addresses are passed through 4 phases in sequence. Each phase attempts a different parsing strategy; addresses that fail one phase are passed to the next:

| Phase | Strategy |
|---|---|
| **Phase 1** | Parses comma-separated addresses (`Street, City, State ZIP`) |
| **Phase 2** | Re-attempts comma-separated parsing with relaxed rules |
| **Phase 3** | Parses space-separated addresses without commas |
| **Phase 4** | Final fallback — attempts positional parsing on remaining addresses |

Addresses that cannot be parsed by any phase are still included in the output with blank city/state/ZIP fields.

---

## Output

One Excel file is saved per input file in the output folder:

```
normalized_geocoded_<original_filename>.xlsx
```

### Output Columns

| Column | Description |
|---|---|
| `Property Address` | Parsed street address |
| `Property City` | Parsed city |
| `Property State` | Parsed 2-letter state code |
| `Date` | Standardized date in `MM/DD/YYYY` format |
| `Revenue` | Cleaned numeric revenue value |
| `Property Zip Code` | 5-digit ZIP code (parsed or geocoded) |

Rows are sorted by `Property Zip Code` descending, with missing ZIPs at the bottom.

---

## Date Formats Supported

The following input formats are automatically detected and converted to `MM/DD/YYYY`:

- `MM/DD/YY`
- `MM/DD/YYYY`
- `YYYY-MM-DD`
- `DD/MM/YY`
- `DD/MM/YYYY`
- Excel serial date numbers

Unrecognized dates are stored as `Invalid Date`.

---

## Geocoding Details

After normalization, rows still missing a valid ZIP code are geocoded if their address contains a comma (indicating enough components to query). The geocoder:

- Uses **Nominatim (OpenStreetMap)** — free, no API key required
- Waits **1 second between requests** to respect rate limits
- Retries up to **3 times** on connection or timeout errors with exponential backoff
- Returns only the **5-digit ZIP** (strips ZIP+4 suffixes)

A summary is printed showing how many ZIPs existed before and after geocoding.

---

## Terminal Output Example

```
Processing file: sold_properties.csv
Phase 1 - Properties normalized: 3420
Phase 1 - Properties kept un-normalized: 580
Phase 2 - Properties normalized: 310
Phase 2 - Properties kept un-normalized: 270
Phase 3 - Properties normalized: 190
Phase 3 - Properties kept un-normalized: 80
Phase 4 - Properties normalized: 45
Phase 4 - Properties kept un-normalized: 35
Total properties processed: 4000
Geocoding data...
Rows needing geocoding: 35
ZIP codes before: 3965, ZIP codes after: 3992
Results saved to: 'normalized_geocoded_sold_properties.xlsx'
Execution time: 87.42 seconds
```

---

## Notes

- All input files in the folder are processed in sequence — no file selection needed
- The script does not modify the original input files
- Geocoding can be slow for large batches of missing ZIPs; plan accordingly
- Files with unsupported extensions are skipped automatically
