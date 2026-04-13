# Sold Properties Processor — `format.py`

Reads a new sold properties CSV, cross-references it against existing Excel records on Google Drive to remove already-processed entries, and saves a deduplicated, formatted Excel output.

---

## What It Does

1. Reads the first `.csv` file found in the input folder
2. Scans a Google Drive folder for existing Excel files and collects all known `property_id` values
3. Removes duplicates within the new CSV and strips any IDs already present in Google Drive
4. Converts money and date columns to proper data types
5. Saves the result as a formatted `.xlsx` file with currency and date formatting applied

---

## Setup

**Requirements:** Python 3.x with the following packages:

```bash
pip install pandas openpyxl
```

---

## Configuration

Open `format.py` and update the three path variables at the top:

```python
# Folder containing your new CSV file
input_folder = r'C:\path\to\input'

# Google Drive folder with existing Excel records (for cross-reference)
gdrive_folder = r'G:\path\to\sold_reports\2026'

# Full path for the output Excel file
output_path = r'C:\path\to\output\Processed_Properties.xlsx'
```

---

## Usage

```bash
python format.py
```

No interactive input is needed — all configuration is set in the file.

---

## Input File Requirements

- Format: `.csv`
- Must be placed in the `input_folder`
- If multiple CSVs are present, **only the first one found** is processed
- Must contain a `property_id` column for deduplication

### Expected Columns

| Column | Format Applied |
|---|---|
| `property_id` | Used for deduplication |
| `price_sold` | Currency (`$#,##0.00`) |
| `total_value` | Currency (`$#,##0.00`) |
| `sold_date` | Date (`yyyy-mm-dd`) |

All other columns are preserved as-is.

---

## Google Drive Cross-Reference

The script scans all `.xlsx` and `.xls` files in the configured `gdrive_folder` and reads their `property_id` column. Any IDs already present in those files are excluded from the output.

- If the Google Drive folder is not found, the cross-reference step is **skipped with a warning** and the script continues
- Files missing a `property_id` column are skipped without stopping the process

---

## Output

A single Excel file saved to `output_path` with:

- Sheet name: `Cleaned_Data`
- Currency formatting on `price_sold` and `total_value`
- Date formatting on `sold_date`
- No row index column

The script also prints a **repeat percentage** showing how many records from the input were already known:

```
Percentage of repeated/existing values: 12.45%
```

---

## Notes

- If the input CSV is open in another program, the script will exit with a clear error message
- Deduplication within the new file keeps the **first occurrence** of each `property_id`
- The original CSV and existing Google Drive files are never modified
