# Excel & CSV File Merger

![Python](https://img.shields.io/badge/Python-3.11+-3776AB?logo=python&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-3.x-150458?logo=pandas&logoColor=white)
![openpyxl](https://img.shields.io/badge/openpyxl-3.x-1D6F42)
![Tests](https://img.shields.io/badge/tests-34%20passing-brightgreen)
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
├── excel_merger.py        # Main script
├── create_samples.py      # Generates the demo input files
├── requirements.txt       # Python dependencies
├── tests/                 # 26 unit tests
├── merged_master.xlsx     # Generated: merged output (gitignored)
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

## Customizing: the `COLUMN_ALIASES` extension point

This is the most reusable idea in the repo. `COLUMN_ALIASES` is a single
dictionary mapping every known source header onto a canonical name:

```python
COLUMN_ALIASES: dict[str, str] = {
    "full name":     "name",
    "employee name": "name",
    "sales amount":  "revenue",
    "amount":        "revenue",
    "salary":        "revenue",
    "territory":     "region",
    ...
}
```

Teaching the merger a new file format is **one line in this dict** — no other
code changes. Everything downstream (currency normalization, deduplication,
column alignment, the summary sheet) keys off canonical names, so the entire
pipeline picks up the new variant automatically. Schema drift becomes
configuration rather than a code change.

### Two caveats worth knowing

**Aliasing can collide.** Several headers intentionally map to the same
canonical name — `Amount`, `Total` and `Salary` all become `revenue`. If a
*single file* contains two of them, they cannot both be `revenue`. `dedupe_columns()`
renames the second to `revenue_2` and logs a warning naming the file:

```
marketing_leads.csv: duplicate canonical column 'revenue' renamed to 'revenue_2'.
Review COLUMN_ALIASES if these should be merged.
```

Nothing is silently dropped, and nothing crashes — you get both columns plus a
prompt to decide whether that mapping was right for your data. Note that only
the canonical `revenue` column gets currency normalization; `revenue_2` is left
as-is precisely because the tool should not guess which one you meant.

**`"first name" → "name"` is lossy.** If a file has separate `First Name` and
`Last Name` columns, only the first is mapped to `name` and the surname stays
under its own column. That is a deliberate simplification, not a bug — but if
your data splits names that way, either remove that alias or pre-join the two
columns before merging.

---

## Development

```bash
pip install -r requirements.txt
pip install pytest ruff

pytest -q          # 34 tests
ruff check .
```

The suite covers alias normalization, currency parsing, whitespace cleaning,
deduplication and source-file tagging, including a regression test for the
alias-collision crash described above.

---

## Possible extensions

Add per-column type coercion, configurable merge keys for true joins (not just concatenation), a dry-run preview, or output to Parquet/SQL for larger datasets.
