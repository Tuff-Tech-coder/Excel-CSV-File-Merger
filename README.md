# Excel & CSV File Merger

![Python](https://img.shields.io/badge/Python-3.10+-3776AB?logo=python&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-2.x-150458?logo=pandas&logoColor=white)
![openpyxl](https://img.shields.io/badge/openpyxl-3.x-1D6F42)
![License](https://img.shields.io/badge/License-MIT-green)

A data-normalization pipeline that scans a folder for Excel and CSV files and intelligently merges them into a single, clean master workbook. Built specifically for the **messy reality** of business data: mismatched column orders, inconsistent naming, currency stored as text, stray whitespace, blank rows, and duplicates.

---

## Why it's useful

Combining spreadsheets from different teams by hand is slow and error-prone — column names never match, totals are formatted as `"$5,400.00"`, and duplicates creep in. This tool encodes those fixes once and applies them consistently, turning a pile of inconsistent files into one analysis-ready dataset with an audit trail.

---

## Features

- **Auto-discovery** of every `.xlsx`, `.xls`, and `.csv` file in a folder.
- **Column-alias mapping** — a configurable dictionary unifies variants (`"Sales Amount"`, `"Amount"`, `"Salary"` → `revenue`; `"Territory"` → `region`) so files merge on meaning, not exact spelling.
- **Currency normalization** — `"$5,400.00"` becomes `5400.0` for real math.
- **Cleanup pass** — strips cell whitespace and drops entirely empty rows.
- **Graceful column alignment** — files missing columns still merge cleanly, with gaps filled rather than erroring.
- **Configurable deduplication** of identical rows (`--no-dedup` to disable).
- **Source tracking** — a `source_file` column records each row's origin.
- **Polished two-sheet output** — a formatted *Merged Data* sheet (styled header, alternating row shading, auto-fit columns) plus a *Merge Summary* sheet with per-file counts, duplicates removed, and final totals.

---

## Tech stack

`Python` · `pandas` · `openpyxl` · `xlrd` · `argparse` · `logging` · `regex`

---

## Project structure

```
excel-merger/
├── excel_merger.py        # Main script
├── create_samples.py      # Generates the demo input files
├── requirements.txt       # Python dependencies
├── merged_master.xlsx     # Sample output
├── merger.log             # Generated: application log
└── sample_input/
    ├── sales_q1.xlsx          # Standard sales data
    ├── sales_q2.xlsx          # Different column order + extra column + a duplicate
    ├── hr_employees.csv       # Missing columns, different naming
    ├── marketing_leads.csv    # Currency-as-text, whitespace, varied date formats
    └── ops_data.xlsx          # Renamed columns + blank rows
```

---

## Setup

```bash
python -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

---

## Usage

```bash
# Merge the included sample files
python excel_merger.py --input ./sample_input --output merged_master.xlsx

# Merge your own folder, keep duplicates
python excel_merger.py --input /path/to/folder --no-dedup
```

**Options:** `--input` · `--output` · `--no-dedup`

---

## How it works

1. Discovers all supported files in the input folder.
2. Loads each, normalizes column names through the alias map, cleans cells, normalizes currency, drops empty rows, and tags rows with their source file.
3. Concatenates everything (pandas aligns on column name, filling gaps).
4. Optionally deduplicates on all columns except the source tag.
5. Writes a formatted workbook with both the merged data and a summary sheet.

---

## Customizing

Extend the `COLUMN_ALIASES` dictionary in `excel_merger.py` to teach the merger any new column-name variants you encounter — no other code changes needed.

---

## Possible extensions

Add per-column type coercion, configurable merge keys for true joins (not just concatenation), a dry-run preview, or output to Parquet/SQL for larger datasets.
