# Sunflower Skiptrace Tag Analyzer — `Sunflower_tags.py`

Scans property Excel files and analyzes the `TAGS` column to identify properties whose most recent Skiptrace tag is **older than 6 months**, flagging them for follow-up.

---

## What It Does

1. Reads all `.xlsx` / `.xls` files from the input folder
2. Parses `Skiptrace` tags from the `TAGS` column of each file
3. Compares each tag's date against a **6-month rolling cutoff** from today
4. Adds a `Tag_Analysis` column with the result for each property
5. Saves the enriched file to the output folder with a `Processed_` prefix

---

## Setup

**Requirements:** Python 3.x with the following packages:

```bash
pip install pandas openpyxl python-dateutil
```

---

## Configuration

The input and output folder paths are hardcoded at the top of the function. Update them before running:

```python
input_folder  = r"C:\path\to\Sunflower_tags\input"
output_folder = r"C:\path\to\Sunflower_tags\output"
```

Both folders are created automatically if they don't exist.

---

## Usage

```bash
python Sunflower_tags.py
```

No interactive input needed. All files in the input folder are processed automatically.

---

## Required Input Column

| Column | Description |
|---|---|
| `TAGS` | Comma-separated list of tags per property. Must contain `Skiptrace` entries in the format `Skiptrace MonthYYYY` (e.g., `Skiptrace January2024`). |

Files missing the `TAGS` column are **skipped** with a warning.

---

## Skiptrace Tag Format

Tags must follow this exact format to be parsed:

```
Skiptrace<MonthYYYY>
```

**Examples:**
- `Skiptrace January2024`
- `Skiptrace March2025`

Tags that don't match this format are ignored. Multiple tags per cell are supported (comma-separated).

---

## Tag Analysis Logic

Each property's `TAGS` cell is evaluated and assigned one of two values in the new `Tag_Analysis` column:

| Result | Condition |
|---|---|
| `Active` | Has at least one Skiptrace tag dated **within the last 6 months** |
| `Active` | Has **no** Skiptrace tags at all |
| `Active` | `TAGS` cell is empty or not a string |
| `OLDER_THAN_6_MONTHS` | All Skiptrace tags found are **older than 6 months** |

> If **any** Skiptrace tag is recent, the property is immediately marked `Active` — older tags on the same property are ignored.

---

## Output

One file is saved per input file in the output folder:

| Input | Output |
|---|---|
| `Miami_List.xlsx` | `Processed_Miami_List.xlsx` |

The output file contains all original columns plus the new `Tag_Analysis` column appended at the end.

The terminal also prints a per-file count of flagged properties:

```
Processing: Miami_List.xlsx... [DONE] - Identified 142 properties to update.
```

---

## Notes

- The 6-month cutoff is calculated dynamically from today's date each time the script runs
- Month parsing uses full English month names (`January`, `February`, etc.) — abbreviated names will not be recognized
- Properties with unparseable tag formats are treated as `Active` (no flag)
- Original files are never modified — all output goes to the output folder
