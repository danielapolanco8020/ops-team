# 🏠 Address Splitter

A Python utility that automatically parses street addresses in Excel (`.xlsx`) and CSV files, splitting combined address strings into a clean **street address** and a separate **unit/apartment** field.

---

## 📋 Features

- Detects a wide range of unit formats including:
  - Keyword-based: `Apt A11`, `Unit 3B`, `Suite 200`, `Room 7C`, `Bldg C`, `Lot 88A`
  - Hash notation: `#4D`, `#F5`, `#5`
  - Floor references: `2nd Floor`, `3rd Fl`, `Floor 2`, `Fl 3`
  - Comma-separated codes: `123 Main St, 4D`
  - Dash-separated codes: `123 Main St - B2`
  - Bare alphanumeric trailing codes: `4D`, `F5`, `B12`
- Inserts two new columns — **Address Modified** and **Apt/Unit** — directly next to the original `ADDRESS` column
- Highlights the new column headers in **yellow** in Excel files
- Processes entire folders of `.xlsx` and `.csv` files in one run
- Includes a built-in self-test suite (20 test cases)

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

### 1. Run the self-tests

To verify everything is working before using on real data:

```bash
python pdx.py
```

This runs 20 built-in test cases and prints pass/fail results.

### 2. Process a folder of files

Edit the bottom of `pdx.py` and uncomment the last two lines, pointing to your folder:

```python
folder_path = 'your_folder_here'
process_files_in_folder(folder_path)
```

Then run:

```bash
python pdx.py
```

The script will:
1. Scan the folder for all `.xlsx` and `.csv` files
2. Look for a column named `ADDRESS` (case-sensitive)
3. Split each address and insert two new columns
4. Save the file in place (overwrites the original — back up first!)
5. Print a confirmation for each processed file

---

## 📊 Output Example

| ADDRESS | Address Modified | Apt/Unit |
|---|---|---|
| 123 Main St Apt A11 | 123 Main St | Apt A11 |
| 456 Elm Ave #4D | 456 Elm Ave | #4D |
| 789 Oak Rd, 3B | 789 Oak Rd | , 3B |
| 600 Spruce Dr 2nd Floor | 600 Spruce Dr | 2nd Floor |
| 900 Willow Ct | 900 Willow Ct | *(empty)* |

In `.xlsx` files, the **Address Modified** and **Apt/Unit** header cells are highlighted in yellow for easy identification.

---

## 🧠 How It Works

The core logic lives in the `find_unit()` function, which applies a **prioritized list of regex patterns** to each address string. Patterns are ordered from most specific to least specific — so a phrase like `"2nd Floor"` is caught by the ordinal floor rule before the generic keyword rule can misfire.

```
Priority order:
  1. Ordinal floor  →  2nd Floor, 3rd Fl
  2. Keyword floor  →  Floor 2, Fl 3
  3. Keyword + code →  Apt A11, Suite 200, Bldg C, No. 12B
  4. Bare #         →  #4D, #F5
  5. After comma    →  123 Main St, 4D
  6. After dash     →  123 Main St - B2
  7. Trailing code  →  ...F5 (end of string)
```

On a match, everything before the match becomes `Address Modified`, and the match itself (to end of string) becomes `Apt/Unit`.

---

## ⚠️ Notes

- The script **overwrites the original files**. Make a backup of your data before running.
- The `ADDRESS` column name is **case-sensitive**. Files without an `ADDRESS` column are skipped.
- Only `.xlsx` and `.csv` files are processed; other file types in the folder are ignored.

---

## 🧪 Running Tests

The self-test suite is embedded at the bottom of `pdx.py`. Run it directly:

```bash
python pdx.py
```

Expected output:
```
  '123 Main St Apt A11'
  '456 Elm Ave Unit 3B'
  ...
20/20 tests passed.
```

To add your own test cases, extend the `tests` list in the `if __name__ == "__main__"` block.
