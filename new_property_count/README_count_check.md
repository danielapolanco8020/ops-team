# Zero Count Checker — `count_check.py`

Scans a folder of property Excel files, counts rows where any `*count*` column equals zero, groups results by marketing channel and quarter, and generates a summary Excel file plus a bar chart.

---

## What It Does

1. Reads all `.xlsx` / `.xls` files from the input folder
2. Extracts the **date** and **marketing channel** from each filename
3. Identifies all columns containing the word `count` (case-insensitive)
4. Counts rows where **any** of those columns has a value of `0`
5. Saves a summary Excel file with results grouped by month, quarter, and channel
6. Displays a bar chart showing zero-count rows by channel and quarter

---

## Setup

**Requirements:** Python 3.x with the following packages:

```bash
pip install pandas openpyxl matplotlib tqdm
```

---

## Configuration

Open `count_check.py` and update the two variables at the bottom:

```python
input_folder = r"C:\path\to\your\input"
output_file = "output_summary.xlsx"  # Saved in the same directory as the script
```

---

## Usage

```bash
python count_check.py
```

No interactive input needed. A progress bar shows processing status as files are read.

---

## Expected Filename Format

Files must follow this naming convention for metadata to be extracted:

```
YYYY-MM-DD <anything> <Channel>.xlsx
```

**Examples:**
- `2025-03-15 Atlas 10K Cold Calling.xlsx`
- `2025-06-01 Atlas 8K Sms.xlsx`
- `2025-09-20 Atlas 5K Direct Mail.xlsx`

| Element | Pattern | Extracted As |
|---|---|---|
| Date | `YYYY-MM-DD` | Used to compute month and quarter |
| Channel | `Cold Calling`, `Sms`, or `Direct Mail` | Stored as `Channel` |

Files that don't match this pattern are **skipped silently**.

---

## Output

### Excel Summary — `output_summary.xlsx`

| Column | Description |
|---|---|
| `Month` | Numeric month extracted from the filename date |
| `Quarter` | Quarter label (`Q1`–`Q4`) derived from the month |
| `Channel` | Marketing channel (`Cold Calling`, `Sms`, `Direct Mail`) |
| `Zero Rows` | Number of rows with at least one `*count*` column equal to `0` |

### Bar Chart

A grouped bar chart is displayed at runtime showing **Zero Rows** on the Y-axis, **Quarter** on the X-axis, and one bar per marketing channel.

---

## Notes

- Any column whose name contains the word `count` (e.g., `Phone Count`, `Email Count`) is included in the zero-row check
- A row is counted if **any** of its count columns equals zero — not necessarily all of them
- Files missing a recognizable channel in the filename are skipped without stopping the batch
- The output Excel is saved in the **script's working directory**, not the input folder
