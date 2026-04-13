# Property Data Auditor (Local Batch) — `audit_local.py`

Automatically processes all Excel files in a specified input folder, running a full suite of quality-control checks on each file and reporting pass/fail results per check. Unlike `auditV2.py`, this version requires **no interactive input** — all checks run automatically.

---

## How It Differs from `auditV2.py`

| Feature | `auditV2.py` | `audit_local.py` |
|---|---|---|
| File input | Single file, typed at runtime | All `.xlsx`/`.xls` in a folder |
| County validation | User provides valid counties | Prints unique counties (no validation) |
| Property type validation | User provides valid types | Prints unique types (no validation) |
| Expected record count | User provides expected number | Just prints actual count |
| Action plan labels | `Prospect Urgent` / `Prospect High` | `30 DAYS` / `60 DAYS` |
| Phone type columns | Single `PHONE TYPE` column | Supports **multiple** `PHONE TYPE X` columns |
| Phone check trigger | Filename contains `sms` or `cold calling` | Filename contains `sms`, `cc`, or `cold calling` |

---

## Setup

**Requirements:** Python 3.x with the following packages:

```bash
pip install pandas openpyxl
```

---

## Usage

1. Open `audit_local.py` and update the folder path:

```python
folder_path = r"C:\path\to\your\Input"
```

2. Place all `.xlsx` / `.xls` files to audit inside that folder.

3. Run the script:

```bash
python audit_local.py
```

The script will automatically process every Excel file found in the folder, one by one.

---

## Audit Checks

| # | Check | Description |
|---|---|---|
| 1 | **Record Count** | Prints total number of rows in the file |
| 2 | **Duplicates #1** | Duplicate `MAILING ADDRESS` + `MAILING ZIP` combinations |
| 3 | **Duplicates #2** | Duplicate `OWNER FULL NAME` + `ADDRESS` + `ZIP` combinations |
| 4 | **Unique Folio** | Same `FOLIO` appearing with different owner/address info |
| 5 | **Owner Full Name value** | Names containing invalid keywords (e.g., `Bank`, `Church`, `County`, `Unknown`, `Not Provided`) |
| 6 | **Owner Full Name complete** | Empty cells in `OWNER FULL NAME` |
| 7 | **Owner Last Name complete** | Empty cells in `OWNER LAST NAME` |
| 8 | **Address complete** | Empty cells in `ADDRESS` |
| 9 | **ZIP complete** | Empty cells in `ZIP` |
| 10 | **County review** | Prints all unique county values for manual review |
| 11 | **Mailing Address complete** | Empty cells in `MAILING ADDRESS` |
| 12 | **Mailing ZIP complete** | Empty cells in `MAILING ZIP` |
| 13 | **Urgent score** | `ACTION PLANS = 30 DAYS` records with `SCORE < 746` |
| 14 | **High score** | `ACTION PLANS = 60 DAYS` records with `SCORE < 545` |
| 15 | **Property status** | Empty cells in `PROPERTY STATUS` |
| 16 | **Property type review** | Prints all unique property type values for manual review |
| 17 | **Tags review** | Records containing unwanted tags: `Liti`, `DNC`, `donotmail`, `Takeoff`, `Undeli`, `Return`, `Dead`, `Do Not Mail`, `Dono`, `Do no`, `Available` |
| 18 | **Absentee address** | Absentee-flagged properties (`ABSENTEE >= 1`) where `ADDRESS` matches `MAILING ADDRESS` |
| 19 | **Phone type — Cold Calling / CC** | *(Filename contains `cold calling` or `cc`)* All `PHONE TYPE` columns checked for `Void`, `Null`, or `Failed` |
| 20 | **Phone type — SMS** | *(Filename contains `sms`)* All `PHONE TYPE` columns checked for `Void`, `Null`, `Failed`, or `Landline` |

---

## Phone Count Check

After the audit, the script runs an additional **phone number summary** for SMS and Cold Calling files:

| Metric | Description |
|---|---|
| Properties with Active Phone Numbers | Rows with data in any `PHONE TYPE` or `PHONE NUMBER` column |
| Properties without Active Phone Numbers | Rows with all phone columns empty |
| Properties without `SKIPTRACE` | Subset of the above that also lack a `skip` tag |

This check only runs if the filename contains `sms`, `cc`, or `cold calling` (case-insensitive).

---

## Required Columns

The Excel files must contain these columns:

`FOLIO`, `OWNER FULL NAME`, `OWNER LAST NAME`, `ADDRESS`, `ZIP`, `COUNTY`, `MAILING ADDRESS`, `MAILING ZIP`, `ACTION PLANS`, `SCORE`, `PROPERTY STATUS`, `PROPERTY TYPE`, `TAGS`, `ABSENTEE`

For Cold Calling or SMS files, also expected: one or more columns starting with `PHONE TYPE` and/or `PHONE NUMBER`

---

## Folder Structure

```
project/
│
├── audit_local.py
│
└── Input/
    ├── Miami_SMS_5000.xlsx
    ├── Orlando_CC_3200.xlsx
    └── Tampa_Mail_4100.xlsx
```

---

## Notes

- County and property type checks print unique values instead of validating — review the output manually to spot out-of-buybox entries
- Multiple `PHONE TYPE` columns (e.g., `PHONE TYPE 1`, `PHONE TYPE 2`) are all checked automatically
- The `cc` keyword trigger for Cold Calling checks matches whole words only, avoiding false positives
- Files that fail to load are reported and skipped without stopping the rest of the batch
