# Property List Splitter — `splitV2.py`

Splits property Excel files into multiple output templates using 6 different splitting strategies, each designed for a different use case. All splits are saved as separate sheets within a single output Excel file per input file.

---

## What It Does

1. Reads all `.xlsx` files from the input folder
2. Displays a summary of all available splitting strategies
3. Prompts the user to select a client split mode (1–6 or `all`)
4. Collects any additional parameters needed for the selected mode
5. Applies the split and saves each result as a separate sheet in an output Excel file

---

## Setup

**Requirements:** Python 3.x with the following packages:

```bash
pip install pandas numpy openpyxl
```

---

## Configuration

Open `splitV2.py` and update the two folder paths at the top:

```python
INPUT_FOLDER  = r"C:\path\to\input"
OUTPUT_FOLDER = r"C:\path\to\output"
```

---

## Usage

```bash
python splitV2.py
```

The script will prompt you to select a client mode and provide any required parameters interactively.

---

## Required Input Columns

All files must contain at minimum:

| Column | Required By |
|---|---|
| `FOLIO` | All clients |
| `PROPERTY TYPE` | All clients |
| `ACTION PLANS` | Clients 1, 3, 6 |
| `SCORE` | Client 3 |
| `COUNTY` | Client 5 |

> Files missing required columns for the selected client are skipped with an error message.

> **Minimum record requirement:** All clients require at least **1,000 records** (Clients 4 and 5 groups require at least **10 records** per group).

---

## Split Modes

### Client 1 — Proportional Split
Splits all records into **Template A** and **Template B**, maintaining a balanced distribution of `ACTION PLANS` across both templates.

- Default split: **50/50**
- Custom split: user enters proportions that must sum to 100 (e.g., `60, 40`)
- Valid `ACTION PLANS` values: `30 DAYS`, `60 DAYS`, `60 DAYS B`, `90 DAYS`, `90 DAYS B`, `90 DAYS C`
- Output sheets: `client_1_template_a`, `client_1_template_b`

---

### Client 2 — Odd/Even Split
Splits records into two groups based on their **row position** in the input file.

- Even-indexed rows → Template Even
- Odd-indexed rows → Template Odd
- No additional parameters needed
- Output sheets: `client_2_even`, `client_2_odd`

---

### Client 3 — Top X + Balanced Split
Preserves the **top X properties by SCORE** in both templates, then splits the remaining records evenly with balanced `ACTION PLANS`.

- User specifies how many top records to preserve (1 to total records)
- Top X records appear in **both** templates
- Remaining records are split 50/50 with balanced action plan distribution
- Valid `ACTION PLANS` values: same as Client 1
- Output sheets: `client_3_template_a`, `client_3_template_b`

---

### Client 4 — Property Type Split
Creates one output sheet per **property type** or **group of property types**.

Two modes:
- **Individual:** each property type gets its own sheet
- **Grouped:** multiple property types are combined into named groups (semicolon-separated)

- Input example (grouped): `Group 1: Residential, Condo; Group 2: Commercial`
- Property type values are **case-sensitive** and must match the data exactly
- Output sheets: `client_4_<property_type>` or `client_4_group_1`, `client_4_group_2`, etc.

---

### Client 5 — County Split
Creates one output sheet per **county** or **group of counties**.

Two modes:
- **Individual:** each county gets its own sheet
- **Grouped:** multiple counties are combined into named groups (semicolon-separated)

- Input example (grouped): `Group 1: Wayne, Macomb; Group 2: Oakland`
- County values must match the data exactly
- Output sheets: `client_5_<county>` or `client_5_group_1`, `client_5_group_2`, etc.

---

### Client 6 — Action Plan Split
Splits records into separate sheets based on `ACTION PLANS` containing `30 DAYS`, `60 DAYS`, or `90 DAYS`.

- Uses partial matching (`str.contains`) — catches variants like `60 DAYS B`, `90 DAYS C`
- Records not matching any keyword are excluded with a warning
- No additional parameters needed
- Output sheets: `client_6_30_days`, `client_6_60_days`, `client_6_90_days`

---

## Output

One Excel file is saved per input file in the output folder:

| Selection | Output Filename |
|---|---|
| Single client | `output_<filename>_client_<N>.xlsx` |
| All clients | `output_<filename>.xlsx` |

Each sheet in the output file corresponds to one template from the selected split.

---

## Notes

- Sheet names are capped at **31 characters** to comply with Excel's limit
- Original row order is preserved within each output template (sorted by original index)
- Duplicate property types or counties across groups are rejected with an error
- Groups with fewer than **10 records** are skipped with a warning
- Running `all` processes all 6 split modes and writes every template into a single output file
