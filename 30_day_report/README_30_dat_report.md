# 30-Day Property Report Processor

Processes Atlas Residential property Excel files and consolidates records flagged with a **30 DAYS** action plan into a single summary report.

---

## What It Does

1. Scans a folder for `.xlsx` / `.xls` files
2. Filters rows where `ACTION PLANS` equals `30 DAYS`
3. Extracts the **date** and **channel** from each filename
4. Converts indicator columns to binary values (`1` or `0`)
5. Outputs a consolidated `Final_Property_Summary.xlsx`

---

## Expected Filename Format

Files must follow this naming convention:

```
YYYY-MM-DD ATLASRESIDENTIAL <size> <channel>.xlsx
```

**Example:** `2025-11-10 ATLASRESIDENTIAL 16K Sms.xlsx`

| Part | Extracted As | Example |
|---|---|---|
| First segment | `DATE` | `2025-11-10` |
| Last segment | `CHANNEL` | `SMS` |

---

## Output Columns

The output file includes the following columns:

| Column | Description |
|---|---|
| `FOLIO` | Property folio ID |
| `APN ADDRESS` | Property address |
| `CITY`, `STATE`, `ZIP`, `COUNTY` | Location fields |
| `ACTION PLANS` | Always `30 DAYS` in this report |
| `PROPERTY STATUS` | Current status of the property |
| `SCORE`, `LIKELY DEAL SCORE`, `BUYBOX SCORE` | Scoring metrics |
| `PROPERTY TYPE` | Type of property |
| `VALUE` | Estimated property value |
| `LINK PROPERTIES` | Related property links |
| `TAGS` | Property tags |
| `DATE` | Extracted from filename |
| `CHANNEL` | Extracted from filename (e.g. `SMS`, `EMAIL`) |
| *(Binary flags)* | See section below |

### Binary Flag Columns

These columns are normalized to `1` (present) or `0` (absent):

`HIDDENGEMS`, `ABSENTEE`, `HIGH EQUITY`, `DOWNSIZING`, `PRE-FORECLOSURE`, `VACANT`, `55+`, `ESTATE`, `INTER FAMILY TRANSFER`, `DIVORCE`, `TAXES`, `PROBATE`, `LOW CREDIT`, `CODE VIOLATIONS`, `BANKRUPTCY`, `LIENS CITY/COUNTY`, `LIENS OTHER`, `LIENS UTILITY`, `LIENS HOA`, `LIENS MECHANIC`, `POOR CONDITION`, `EVICTION`, `30-60 DAYS`, `JUDGEMENT`, `DEBT COLLECTION`, `DEFAULT RISK`

---

## Setup

**Requirements:** Python 3.x with the following packages:

```bash
pip install pandas openpyxl
```

---

## Usage

1. Open `30_day_report.py`
2. Update the path at the bottom of the file to point to your folder:

```python
process_atlas_files(r"C:\path\to\your\folder")
```

3. Run the script:

```bash
python 30_day_report.py
```

4. The output file `Final_Property_Summary.xlsx` will be created in the **same directory where you run the script**.

---

## Notes

- Files with no rows matching `30 DAYS` are silently skipped
- Binary flag columns missing from a source file default to `0`
- All `ACTION PLANS` comparisons are case-insensitive
