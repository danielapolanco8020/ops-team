# 🗂️ Marketing Cadence Consolidator

A Python utility that merges multiple Excel files from real estate marketing campaigns into a single consolidated file per cadence — with column validation, selective processing, and a clean summary report.

---

## 📋 Features

- Interactive prompt at runtime to choose which cadence to consolidate
- Validates required columns per cadence before merging — skips invalid files with a clear warning
- Reports optional/extra columns found in each file
- Merges all valid files for the selected cadence into one output file
- Sorts consolidated data by lead scores (`BUYBOX SCORE`, `LIKELY DEAL SCORE`, `SCORE`) descending
- Auto-updates the output filename with the new consolidated record count (e.g. `9.5K`)
- Prints a final summary with total records per cadence

---

## 📂 Supported Cadences

| Code | Cadence | File keyword match |
|---|---|---|
| `dm` | Direct Mail | `*Direct Mail*.xlsx` |
| `cc` | Cold Calling | `*Cold Calling*.xlsx` |
| `sms` | SMS | `*SMS*.xlsx` |
| `all` | All three | Runs all of the above |

---

## ⚙️ Requirements

- Python 3.7+
- pandas
- openpyxl

Install dependencies:

```bash
pip install pandas openpyxl
```

---

## 🚀 Usage

### 1. Prepare your folder structure

Place all input Excel files inside a `Processed_Files` folder in the same directory as the script:

```
project/
├── consolidate.py
└── Processed_Files/
    ├── Batch 1 2K Direct Mail.xlsx
    ├── Batch 2 3.5K Direct Mail.xlsx
    ├── Batch 1 1K Cold Calling.xlsx
    └── Batch 1 2K SMS.xlsx
```

### 2. Run the script

```bash
python consolidate.py
```

### 3. Select a cadence at the prompt

```
==================================================
  CONSOLIDATION TOOL
==================================================
Which cadence would you like to consolidate?
  dm   → Direct Mail
  cc   → Cold Calling
  sms  → SMS
  all  → All three
--------------------------------------------------
Enter cadence code: dm
```

### 4. Review console output and find results

Consolidated files are saved to the `Consolidated_Files` folder, which is created automatically if it doesn't exist.

---

## 🔍 Column Validation

Before merging, every file is checked against the required column list for its cadence. The console will report:

```
  Checking: Batch 1 2K Direct Mail.xlsx
     ❌ MISSING required columns:
        - TARGETED POSTCARD
     ⚠️  Skipped due to missing required columns.

  Checking: Batch 2 3.5K Direct Mail.xlsx
     ℹ️  Optional/extra columns found:
        + NOTES
     ✅ Passed validation — 3,500 rows
```

Files with missing **required** columns are skipped. Files with extra or optional columns are accepted and the differences are logged.

---

## 📁 Required Columns per Cadence

### Direct Mail (`dm`)
`FOLIO` `APN` `OWNER FULL NAME` `OWNER FIRST NAME` `OWNER LAST NAME` `ADDRESS` `CITY` `STATE` `ZIP` `COUNTY` `MAILING ADDRESS` `MAILING CITY` `MAILING STATE` `MAILING ZIP` `GOLDEN ADDRESS` `GOLDEN CITY` `GOLDEN STATE` `GOLDEN ZIP CODE` `ACTION PLANS` `PROPERTY STATUS` `SCORE` `LIKELY DEAL SCORE` `BUYBOX SCORE` `PROPERTY TYPE` `VALUE` `LINK PROPERTIES` `TAGS` `HIDDENGEMS` `ABSENTEE` `HIGH EQUITY` `DOWNSIZING` `PRE-FORECLOSURE` `VACANT` `55+` `ESTATE` `INTER FAMILY TRANSFER` `DIVORCE` `TAXES` `PROBATE` `LOW CREDIT` `CODE VIOLATIONS` `BANKRUPTCY` `LIENS CITY/COUNTY` `LIENS OTHER` `LIENS UTILITY` `LIENS HOA` `LIENS MECHANIC` `POOR CONDITION` `EVICTION` `30-60 DAYS` `JUDGEMENT` `DEBT COLLECTION` `DEFAULT RISK` `MARKETING DM COUNT` `ESTIMATED CASH OFFER` `MAIN DISTRESS #1` `MAIN DISTRESS #2` `MAIN DISTRESS #3` `MAIN DISTRESS #4` `TARGETED MESSAGE #1` `TARGETED MESSAGE #2` `TARGETED MESSAGE #3` `TARGETED MESSAGE #4` `TARGETED GROUP NAME` `TARGETED GROUP MESSAGE` `TARGETED POSTCARD`

### Cold Calling (`cc`)
`FOLIO` `APN` `OWNER FULL NAME` `OWNER FIRST NAME` `OWNER LAST NAME` `SECOND OWNER FULL NAME` `ADDRESS` `CITY` `STATE` `ZIP` `COUNTY` `MAILING ADDRESS` `MAILING CITY` `MAILING STATE` `MAILING ZIP` `GOLDEN ADDRESS` `GOLDEN CITY` `GOLDEN STATE` `GOLDEN ZIP CODE` `ACTION PLANS` `PROPERTY STATUS` `SCORE` `LIKELY DEAL SCORE` `BUYBOX SCORE` `PROPERTY TYPE` `VALUE` `LINK PROPERTIES` `TAGS` `HIDDENGEMS` `ABSENTEE` `HIGH EQUITY` `DOWNSIZING` `PRE-FORECLOSURE` `VACANT` `55+` `ESTATE` `INTER FAMILY TRANSFER` `DIVORCE` `TAXES` `PROBATE` `LOW CREDIT` `CODE VIOLATIONS` `BANKRUPTCY` `LIENS CITY/COUNTY` `LIENS OTHER` `LIENS UTILITY` `LIENS HOA` `LIENS MECHANIC` `POOR CONDITION` `EVICTION` `30-60 DAYS` `JUDGEMENT` `DEBT COLLECTION` `DEFAULT RISK` `MARKETING CC COUNT`

### SMS (`sms`)
`FOLIO` `APN` `OWNER FULL NAME` `OWNER FIRST NAME` `OWNER LAST NAME` `ADDRESS` `CITY` `STATE` `ZIP` `COUNTY` `MAILING ADDRESS` `MAILING CITY` `MAILING STATE` `MAILING ZIP` `GOLDEN ADDRESS` `GOLDEN CITY` `GOLDEN STATE` `GOLDEN ZIP CODE` `ACTION PLANS` `PROPERTY STATUS` `SCORE` `LIKELY DEAL SCORE` `BUYBOX SCORE` `PROPERTY TYPE` `VALUE` `LINK PROPERTIES` `TAGS` `HIDDENGEMS` `ABSENTEE` `HIGH EQUITY` `DOWNSIZING` `PRE-FORECLOSURE` `VACANT` `55+` `ESTATE` `INTER FAMILY TRANSFER` `DIVORCE` `TAXES` `PROBATE` `LOW CREDIT` `CODE VIOLATIONS` `BANKRUPTCY` `LIENS CITY/COUNTY` `LIENS OTHER` `LIENS UTILITY` `LIENS HOA` `LIENS MECHANIC` `POOR CONDITION` `EVICTION` `30-60 DAYS` `JUDGEMENT` `DEBT COLLECTION` `DEFAULT RISK` `MARKETING SMS COUNT` `MAIN DISTRESS #1` `MAIN DISTRESS #2` `MAIN DISTRESS #3` `MAIN DISTRESS #4` `TARGETED MESSAGE #1` `TARGETED MESSAGE #2` `TARGETED MESSAGE #3` `TARGETED MESSAGE #4` `TARGETED GROUP NAME` `TARGETED GROUP MESSAGE`

---

## 📊 Console Output Example

```
==================================================
  CONSOLIDATION TOOL
==================================================
Which cadence would you like to consolidate?
  dm   → Direct Mail
  cc   → Cold Calling
  sms  → SMS
  all  → All three
--------------------------------------------------
Enter cadence code: all

  ✅ Running: Direct Mail, Cold Calling, SMS
==================================================

Processing 'Direct Mail'... (2 file(s) found)
  Checking: Batch 1 2K Direct Mail.xlsx
     ✅ Passed validation — 2,000 rows
  Checking: Batch 2 3.5K Direct Mail.xlsx
     ✅ Passed validation — 3,500 rows

  -> ✅ Saved: Batch 1 5.5K Direct Mail.xlsx (5,500 total rows)

========================================
FINAL SUMMARY
========================================
Direct Mail    : 5,500 records
Cold Calling   : 3,200 records
SMS            : 2,100 records
----------------------------------------
TOTAL          : 10,800 records
========================================
```

---

## ⚠️ Notes

- Input files must contain the cadence keyword in their filename (`Direct Mail`, `Cold Calling`, or `SMS`) to be detected.
- The output filename is derived from the **first matched file** in the folder — it must already contain a K-formatted number (e.g. `2K` or `3.5K`) surrounded by spaces for the rename to work correctly.
- Output files are saved to `Consolidated_Files/` and will **overwrite** any existing file with the same name.
- If all files in a cadence fail validation, that cadence is skipped entirely — no partial output is written.
