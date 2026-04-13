# 📅 Direct Mail Weekly Splitter

A Python utility that takes consolidated Direct Mail Excel files and distributes their records evenly across 4 weekly tabs — ready for week-by-week outreach campaigns.

---

## 📋 Features

- Automatically finds all Direct Mail Excel files in a target folder
- Splits total records into 4 equal weekly chunks
- Handles uneven row counts gracefully — no records are lost or duplicated
- Writes `Week 1` through `Week 4` sheets directly into the original file
- Preserves the original data sheet untouched
- Overwrites existing Week tabs if the file has been processed before

---

## 💡 How the Split Works

Records are divided in the order they appear in the file. If you run this after the [Consolidation step](#pipeline-context), records will already be sorted by lead score descending — meaning **Week 1 receives the highest-scoring leads** and Week 4 the lowest.

For uneven totals, `numpy.array_split` distributes the remainder across the first chunks:

| Total Rows | Week 1 | Week 2 | Week 3 | Week 4 |
|---|---|---|---|---|
| 100 | 25 | 25 | 25 | 25 |
| 101 | 26 | 25 | 25 | 25 |
| 103 | 26 | 26 | 26 | 25 |

---

## ⚙️ Requirements

- Python 3.7+
- pandas
- numpy
- openpyxl

Install dependencies:

```bash
pip install pandas numpy openpyxl
```

---

## 🚀 Usage

### 1. Prepare your folder

Place all Direct Mail Excel files inside a `Processed_Files` folder in the same directory as the script. Files must contain `Direct Mail` in the filename to be detected:

```
project/
├── moss.py
└── Processed_Files/
    ├── Batch 1 5.5K Direct Mail.xlsx
    └── Batch 2 3K Direct Mail.xlsx
```

### 2. Run the script

```bash
python moss.py
```

### 3. Check your files

Each matched file will now contain its original sheet plus four new weekly tabs:

```
Batch 1 5.5K Direct Mail.xlsx
├── Sheet1          ← original data, untouched
├── Week 1          ← highest-scoring ~25% of records
├── Week 2
├── Week 3
└── Week 4          ← lowest-scoring ~25% of records
```

---

## 📊 Console Output Example

```
Searching for 'Direct Mail' files in: /path/to/Processed_Files

--- Processing file: Batch 1 5.5K Direct Mail.xlsx ---
Total rows found: 5,500
  -> Generated: Week 1 with 1,375 rows.
  -> Generated: Week 2 with 1,375 rows.
  -> Generated: Week 3 with 1,375 rows.
  -> Generated: Week 4 with 1,375 rows.
Success: File updated with 4 weeks.

--- Processing file: Batch 2 3K Direct Mail.xlsx ---
Total rows found: 3,001
  -> Generated: Week 1 with 751 rows.
  -> Generated: Week 2 with 750 rows.
  -> Generated: Week 3 with 750 rows.
  -> Generated: Week 4 with 750 rows.
Success: File updated with 4 weeks.
```

---

## 🔁 Pipeline Context

This script is the final step in a 3-stage pipeline:

```
1. pdx.py             →  Split raw addresses into street + unit columns
2. consolidate.py     →  Merge batches by cadence, validate columns, sort by score
3. moss.py            →  Split consolidated Direct Mail file into 4 weekly tabs
```

Running them in order ensures the weekly split is always working from clean, validated, score-sorted data.

---

## ⚠️ Notes

- The script **modifies the original files in place**. Make a backup before running.
- Only `.xlsx` and `.xls` files containing `Direct Mail` in the filename are processed.
- The split is purely positional — row order in the file determines which week each record lands in.
- If `Week 1`–`Week 4` sheets already exist in a file, they will be overwritten.
- Only the **first sheet** of each file is read and split.
