# Dataflick File Formatter — `dataflick_format.py`

Reads property Excel files, keeps only the columns required for Dataflick upload, and splits large files into chunks of 20,000 rows. Output files are saved as new Excel files.

---

## What It Does

1. Reads all `.xlsx` files from the input folder
2. Retains only the 16 columns needed for Dataflick
3. If a file has more than 20,000 rows, splits it into numbered chunks
4. Saves each output file to the output folder

---

## Setup

```bash
pip install pandas openpyxl
```

---

## Configuration

Update the two folder paths at the bottom of the script:

```python
input_folder_path  = r"C:\path\to\input_dataflick"
output_folder_path = r"C:\path\to\output_dataflick"
```

The output folder is created automatically if it doesn't exist.

---

## Usage

```bash
python dataflick_format.py
```

No interactive input needed.

---

## Required Input Columns

The input files must contain these exact columns:

| Column | Description |
|---|---|
| `property_address` | Property street address |
| `property_city` | Property city |
| `property_state` | Property state |
| `property_zip` | Property ZIP code |
| `mailing_address` | Mailing street address |
| `mailing_address_city` | Mailing city |
| `mailing_address_state` | Mailing state |
| `mailing_address_zip` | Mailing ZIP code |
| `phone_1` | Primary phone number |
| `phone_1_type` | Type of primary phone |
| `phone_2` | Secondary phone number |
| `phone_2_type` | Type of secondary phone |
| `phone_3` | Third phone number |
| `phone_3_type` | Type of third phone |
| `phone_4` | Fourth phone number |
| `phone_4_type` | Type of fourth phone |

---

## Output

| Condition | Output Filename |
|---|---|
| File ≤ 20,000 rows | `<original_name>_output.xlsx` |
| File > 20,000 rows | `<original_name>_chunk_1.xlsx`, `<original_name>_chunk_2.xlsx`, etc. |

---

## Notes

- Only `.xlsx` files are processed — other formats are ignored
- The chunk size of 20,000 rows is hardcoded but can be changed via the `chunk_size` parameter in `process_files()`
- Original files are never modified
