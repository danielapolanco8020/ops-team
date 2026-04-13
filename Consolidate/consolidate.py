import os
import pandas as pd
import glob
import re

input_folder  = 'Processed_Files'
output_folder = 'Consolidated_Files'

# ── Required columns per cadence ──────────────────────────────────────────────
REQUIRED_COLUMNS = {
    "Direct Mail": [
        "FOLIO","APN","OWNER FULL NAME","OWNER FIRST NAME","OWNER LAST NAME",
        "ADDRESS","CITY","STATE","ZIP","COUNTY",
        "MAILING ADDRESS","MAILING CITY","MAILING STATE","MAILING ZIP",
        "GOLDEN ADDRESS","GOLDEN CITY","GOLDEN STATE","GOLDEN ZIP CODE",
        "ACTION PLANS","PROPERTY STATUS","SCORE","LIKELY DEAL SCORE","BUYBOX SCORE",
        "PROPERTY TYPE","VALUE","LINK PROPERTIES","TAGS","HIDDENGEMS","ABSENTEE",
        "HIGH EQUITY","DOWNSIZING","PRE-FORECLOSURE","VACANT","55+","ESTATE",
        "INTER FAMILY TRANSFER","DIVORCE","TAXES","PROBATE","LOW CREDIT",
        "CODE VIOLATIONS","BANKRUPTCY","LIENS CITY/COUNTY","LIENS OTHER",
        "LIENS UTILITY","LIENS HOA","LIENS MECHANIC","POOR CONDITION","EVICTION",
        "30-60 DAYS","JUDGEMENT","DEBT COLLECTION","DEFAULT RISK",
        "MARKETING DM COUNT","ESTIMATED CASH OFFER",
        "MAIN DISTRESS #1","MAIN DISTRESS #2","MAIN DISTRESS #3","MAIN DISTRESS #4",
        "TARGETED MESSAGE #1","TARGETED MESSAGE #2","TARGETED MESSAGE #3","TARGETED MESSAGE #4",
        "TARGETED GROUP NAME","TARGETED GROUP MESSAGE","TARGETED POSTCARD",
    ],
    "Cold Calling": [
        "FOLIO","APN","OWNER FULL NAME","OWNER FIRST NAME","OWNER LAST NAME",
        "SECOND OWNER FULL NAME",
        "ADDRESS","CITY","STATE","ZIP","COUNTY",
        "MAILING ADDRESS","MAILING CITY","MAILING STATE","MAILING ZIP",
        "GOLDEN ADDRESS","GOLDEN CITY","GOLDEN STATE","GOLDEN ZIP CODE",
        "ACTION PLANS","PROPERTY STATUS","SCORE","LIKELY DEAL SCORE","BUYBOX SCORE",
        "PROPERTY TYPE","VALUE","LINK PROPERTIES","TAGS","HIDDENGEMS","ABSENTEE",
        "HIGH EQUITY","DOWNSIZING","PRE-FORECLOSURE","VACANT","55+","ESTATE",
        "INTER FAMILY TRANSFER","DIVORCE","TAXES","PROBATE","LOW CREDIT",
        "CODE VIOLATIONS","BANKRUPTCY","LIENS CITY/COUNTY","LIENS OTHER",
        "LIENS UTILITY","LIENS HOA","LIENS MECHANIC","POOR CONDITION","EVICTION",
        "30-60 DAYS","JUDGEMENT","DEBT COLLECTION","DEFAULT RISK",
        "MARKETING CC COUNT",
    ],
    "SMS": [
        "FOLIO","APN","OWNER FULL NAME","OWNER FIRST NAME","OWNER LAST NAME",
        "ADDRESS","CITY","STATE","ZIP","COUNTY",
        "MAILING ADDRESS","MAILING CITY","MAILING STATE","MAILING ZIP",
        "GOLDEN ADDRESS","GOLDEN CITY","GOLDEN STATE","GOLDEN ZIP CODE",
        "ACTION PLANS","PROPERTY STATUS","SCORE","LIKELY DEAL SCORE","BUYBOX SCORE",
        "PROPERTY TYPE","VALUE","LINK PROPERTIES","TAGS","HIDDENGEMS","ABSENTEE",
        "HIGH EQUITY","DOWNSIZING","PRE-FORECLOSURE","VACANT","55+","ESTATE",
        "INTER FAMILY TRANSFER","DIVORCE","TAXES","PROBATE","LOW CREDIT",
        "CODE VIOLATIONS","BANKRUPTCY","LIENS CITY/COUNTY","LIENS OTHER",
        "LIENS UTILITY","LIENS HOA","LIENS MECHANIC","POOR CONDITION","EVICTION",
        "30-60 DAYS","JUDGEMENT","DEBT COLLECTION","DEFAULT RISK",
        "MARKETING SMS COUNT",
        "MAIN DISTRESS #1","MAIN DISTRESS #2","MAIN DISTRESS #3","MAIN DISTRESS #4",
        "TARGETED MESSAGE #1","TARGETED MESSAGE #2","TARGETED MESSAGE #3","TARGETED MESSAGE #4",
        "TARGETED GROUP NAME","TARGETED GROUP MESSAGE",
    ],
}

# Short-code → keyword mapping
CADENCE_MAP = {
    "dm":  "Direct Mail",
    "cc":  "Cold Calling",
    "sms": "SMS",
    "all": None,   # special: run all three
}

# ── Helpers ───────────────────────────────────────────────────────────────────
def format_count_to_k(count):
    k_value = count / 1000
    return f"{int(k_value)}K" if k_value == int(k_value) else f"{round(k_value, 1)}K"

def construct_new_filename(base_filename, new_total_k):
    pattern     = r'\s\d+(\.\d+)?K\s'
    replacement = f" {new_total_k} "
    return re.sub(pattern, replacement, base_filename)

def validate_columns(df, filename, keyword):
    """
    Check required columns are present.
    Report any missing required columns and any extra (optional) columns.
    Returns True if all required columns are present, False otherwise.
    """
    required  = set(REQUIRED_COLUMNS[keyword])
    present   = set(df.columns)
    missing   = required - present
    extra     = present - required

    ok = True
    if missing:
        print(f"     ❌ MISSING required columns in '{filename}':")
        for col in sorted(missing):
            print(f"        - {col}")
        ok = False
    if extra:
        print(f"     ℹ️  Optional/extra columns found in '{filename}':")
        for col in sorted(extra):
            print(f"        + {col}")
    return ok

# ── Cadence selection prompt ──────────────────────────────────────────────────
print("=" * 50)
print("  CONSOLIDATION TOOL")
print("=" * 50)
print("Which cadence would you like to consolidate?")
print("  dm   → Direct Mail")
print("  cc   → Cold Calling")
print("  sms  → SMS")
print("  all  → All three")
print("-" * 50)

while True:
    choice = input("Enter cadence code: ").strip().lower()
    if choice in CADENCE_MAP:
        break
    print(f"  ⚠️  '{choice}' is not valid. Please enter dm, cc, sms, or all.")

# Build the list of keywords to process
if choice == "all":
    keywords_to_run = list(REQUIRED_COLUMNS.keys())
else:
    keywords_to_run = [CADENCE_MAP[choice]]

print(f"\n  ✅ Running: {', '.join(keywords_to_run)}")
print("=" * 50 + "\n")

# ── Main processing ───────────────────────────────────────────────────────────
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

print(f"Input folder : {os.path.abspath(input_folder)}")
print("--- Starting Consolidation ---\n")

summary_totals = {}

for keyword in keywords_to_run:
    search_pattern = os.path.join(input_folder, f"*{keyword}*.xlsx")
    files          = glob.glob(search_pattern)

    if not files:
        print(f"⚠️  No Excel files found for: '{keyword}'")
        continue

    print(f"Processing '{keyword}'... ({len(files)} file(s) found)")

    dataframes   = []
    skipped      = 0

    for file in files:
        fname = os.path.basename(file)
        try:
            df = pd.read_excel(file, dtype=str)
            print(f"  Checking: {fname}")
            if validate_columns(df, fname, keyword):
                dataframes.append(df)
                print(f"     ✅ Passed validation — {len(df):,} rows")
            else:
                print(f"     ⚠️  Skipped due to missing required columns.")
                skipped += 1
        except Exception as e:
            print(f"     ❌ Could not read '{fname}': {e}")
            skipped += 1

    if not dataframes:
        print(f"  ❌ No valid files to consolidate for '{keyword}'. Skipping.\n")
        continue

    if skipped:
        print(f"\n  ⚠️  {skipped} file(s) were skipped for '{keyword}'.")

    try:
        consolidated_df = pd.concat(dataframes, ignore_index=True)

        cols_to_sort = ['BUYBOX SCORE', 'LIKELY DEAL SCORE', 'SCORE']
        sort_cols    = [c for c in cols_to_sort if c in consolidated_df.columns]
        for col in sort_cols:
            consolidated_df[col] = pd.to_numeric(consolidated_df[col], errors='coerce')
        if sort_cols:
            consolidated_df = consolidated_df.sort_values(
                by=sort_cols,
                ascending=[False] * len(sort_cols)
            )

        total_rows   = len(consolidated_df)
        summary_totals[keyword] = total_rows
        total_k_str  = format_count_to_k(total_rows)

        first_filename = os.path.basename(files[0])
        new_filename   = construct_new_filename(first_filename, total_k_str)
        output_path    = os.path.join(output_folder, new_filename)

        consolidated_df.to_excel(output_path, index=False)
        print(f"\n  -> ✅ Saved: {new_filename} ({total_rows:,} total rows)\n")

    except Exception as e:
        print(f"  -> ❌ Error consolidating '{keyword}': {e}\n")

# ── Summary ───────────────────────────────────────────────────────────────────
print("=" * 40)
print("FINAL SUMMARY")
print("=" * 40)

total_general = 0
for category, count in summary_totals.items():
    print(f"{category:<15}: {count:,.0f} records")
    total_general += count

print("-" * 40)
print(f"TOTAL          : {total_general:,.0f} records")
print("=" * 40)
