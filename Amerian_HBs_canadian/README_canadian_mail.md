# Canadian Mailing Address Filter — `canadian_mail.py`

Scans one or more CSV files containing property/contact records and extracts all rows where the **mailing address is in Canada**, merging them into a single output CSV.

---

## What It Does

1. Reads all `.csv` files from a specified input folder
2. Auto-detects the delimiter of each file (comma, pipe, tab, etc.)
3. Locates the `MailingCity` and `MailingState` columns (flexible naming)
4. Flags rows where `MailingState` matches a Canadian province or territory code
5. Collects all matching rows and saves them to a single merged CSV
6. Adds a `SourceFile` column to track which file each record came from

---

## Canadian Provinces & Territories Recognized

The script checks for these standard 2-letter codes in the `MailingState` column:

| Code | Province / Territory |
|---|---|
| `AB` | Alberta |
| `BC` | British Columbia |
| `MB` | Manitoba |
| `NB` | New Brunswick |
| `NL` | Newfoundland and Labrador |
| `NS` | Nova Scotia |
| `NT` | Northwest Territories |
| `NU` | Nunavut |
| `ON` | Ontario |
| `PE` | Prince Edward Island |
| `QC` | Quebec |
| `SK` | Saskatchewan |
| `YT` | Yukon |

Matching is **case-insensitive** — `on`, `ON`, and `On` all match.

---

## Setup

**Requirements:** Python 3.x with the following packages:

```bash
pip install pandas
```

---

## Usage

1. Open `canadian_mail.py` and update the paths at the bottom:

```python
input_path = r"C:\path\to\your\input\*.csv"
output_file = r"C:\path\to\your\output\canadian_locations_merged.csv"
```

2. Run the script:

```bash
python canadian_mail.py
```

3. The output CSV is saved to the path specified in `output_file`.

---

## Input File Requirements

- Format: `.csv`
- Must contain a **mailing state/province** column and ideally a **mailing city** column
- The script accepts flexible column naming — any of the following will be recognized:

| Field | Accepted Column Names |
|---|---|
| Mailing State | `MailingState`, `State`, `Province`, `Mailing_State`, `Mailing State`, `mailing-state` |
| Mailing City | `MailingCity`, `City`, `mailing_city`, `Mailing City`, `mailing-city` |

> Files missing both columns are skipped with a warning.

---

## Output

A single merged CSV file containing all Canadian rows from all input files, with an added column:

| Column | Description |
|---|---|
| *(all original columns)* | Preserved as-is from source files |
| `SourceFile` | Name of the file the row came from |

---

## Notes

- The delimiter is auto-detected per file using Python's `csv.Sniffer` — no manual configuration needed
- If delimiter detection fails, the script defaults to comma (`,`)
- Files are inspected (first 5 lines printed) at runtime to aid debugging
- Rows with invalid or unrecognized state values are simply not matched — no rows are dropped
- If no Canadian rows are found across all files, no output file is created
- Temporary helper columns (`IsInCanada`, `MailingStateClean`) are removed before saving
