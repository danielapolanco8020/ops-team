# 📬 DM Count Excel Processor

A Python utility that processes real estate marketing Excel files — adjusting direct mail counts, removing outdated offer columns, and automatically splitting data into organized tabs with calculated cash offer prices.

---

## 📋 Features

- Decrements the `MARKETING DM COUNT` column by 1 to correct off-by-one tracking
- Removes the legacy `ESTIMATED CASH OFFER` column
- Splits data into separate sheets based on DM Count value (0–5 and 6+)
- Calculates two cash offer columns per sheet based on property value:
  - Always includes a **90% offer**
  - Includes a **60% or 65% offer** depending on contact history
- Inserts offer columns directly after the `MARKETING DM COUNT` column
- Processes entire folders of `.xlsx`, `.xls`, and `.xlsm` files in one run
- Skips empty groups silently — no blank sheets created

---

## 💡 Cash Offer Logic

| Sheet | Filter | CASH OFFER 90% | 2nd Offer |
|---|---|---|---|
| DM Count 0 | count == 0 | VALUE × 90% | VALUE × 60% |
| DM Count 1 | count == 1 | VALUE × 90% | VALUE × 60% |
| DM Count 2 | count == 2 | VALUE × 90% | VALUE × 60% |
| DM Count 3 | count == 3 | VALUE × 90% | VALUE × 65% |
| DM Count 4 | count == 4 | VALUE × 90% | VALUE × 65% |
| DM Count 5 | count == 5 | VALUE × 90% | VALUE × 65% |
| DM Count 6 or more | count >= 6 | VALUE × 90% | VALUE × 65% |

Properties with a higher contact history (3+) receive a slightly higher secondary offer rate (65% vs 60%), reflecting increased seller motivation.

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

### 1. Prepare your folder

Place all Excel files to be processed inside a folder named `Processed_Files` in the same directory as the script:

```
project/
├── dm_processor.py
└── Processed_Files/
    ├── leads_batch_01.xlsx
    ├── leads_batch_02.xlsx
    └── ...
```

### 2. Run the script

```bash
python dm_processor.py
```

### 3. Custom folder path

You can pass a different folder path directly in the code:

```python
process_excel_files(folder_path='your_folder_here')
```

---

## 📊 Output Example

Each processed file will have its first sheet updated and new sheets added, one per DM Count group. For example, a file with records spanning counts 0–8 would produce:

| Sheet Name | Contents |
|---|---|
| *(original sheet)* | Updated with DM Count decremented |
| DM Count 0 | Rows where count == 0, with 90% and 60% offer columns |
| DM Count 1 | Rows where count == 1, with 90% and 60% offer columns |
| DM Count 2 | Rows where count == 2, with 90% and 60% offer columns |
| DM Count 3 | Rows where count == 3, with 90% and 65% offer columns |
| DM Count 4 | Rows where count == 4, with 90% and 65% offer columns |
| DM Count 5 | Rows where count == 5, with 90% and 65% offer columns |
| DM Count 6 or more | Rows where count >= 6, with 90% and 65% offer columns |

---

## 📁 Required Column Names

The script expects these exact column names in your Excel file (case-sensitive):

| Column | Description |
|---|---|
| `MARKETING DM COUNT` | Number of times a property has been mailed |
| `VALUE` | Property value used to calculate cash offers |
| `ESTIMATED CASH OFFER` | *(optional)* Legacy column — removed automatically if present |

---

## ⚠️ Notes

- The script **overwrites the original files**. Make a backup of your data before running.
- Only the **first sheet** of each file is read and processed.
- Files without a `MARKETING DM COUNT` column are skipped with an error message.
- Sheets for DM Count groups with no matching rows are not created.
- Supported file types: `.xlsx`, `.xls`, `.xlsm`
