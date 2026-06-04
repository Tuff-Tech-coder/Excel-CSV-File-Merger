# Excel & CSV File Merger

A Python automation script that scans a folder for all Excel and CSV files, intelligently merges them into a single normalized master file, and outputs a clean, formatted Excel workbook. Built for real-world messy data.

## Features

- **Auto-discovery**: finds all `.xlsx`, `.xls`, and `.csv` files in a folder
- **Column normalization**: strips whitespace, lowercases, and maps common aliases (e.g. `"Sales Amount"` → `revenue`, `"Territory"` → `region`)
- **Missing columns**: files without certain columns still merge cleanly — missing values are filled with blanks
- **Duplicate removal**: deduplicates identical rows across files (configurable)
- **Inconsistent formatting**: normalizes currency strings (`"$5,400.00"` → `5400.0`), strips cell whitespace, drops empty rows
- **Source tracking**: adds a `source_file` column to every row
- **Formatted output**: styled Excel file with alternating row shading, bold headers, auto-fit columns, and a summary statistics sheet

## Project Structure

```
excel-merger/
├── excel_merger.py        # Main script
├── create_samples.py      # Script that generated the sample input files
├── requirements.txt       # Python dependencies
├── merged_master.xlsx     # Sample output (28 rows from 5 messy files)
├── merger.log             # Application log
└── sample_input/
    ├── sales_q1.xlsx          # Standard sales data
    ├── sales_q2.xlsx          # Different column order + extra column + duplicate
    ├── hr_employees.csv       # Missing revenue/date, different column names
    ├── marketing_leads.csv    # Inconsistent formatting ($, whitespace, date formats)
    └── ops_data.xlsx          # Renamed columns + blank rows
```

## Setup

```bash
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## Usage

```bash
# Merge all Excel/CSV files in the sample_input folder
python excel_merger.py

# Specify a custom input folder and output file
python excel_merger.py --input /path/to/my/data --output results.xlsx

# Skip duplicate removal
python excel_merger.py --no-dedup
```

## Command-Line Options

| Option | Default | Description |
|---|---|---|
| `--input` | `./sample_input` | Folder to scan for files |
| `--output` | `merged_master.xlsx` | Output Excel filename |
| `--no-dedup` | off | Disable duplicate row removal |

## Sample Output (`merged_master.xlsx`)

The **"Merged Data"** sheet contains all rows from all 5 files, aligned by canonical column name:

| name | email | revenue | region | date | department | notes | hire date | salary | source_file |
|---|---|---|---|---|---|---|---|---|---|
| Alice Johnson | alice@corp.com | 12500.0 | North | 2024-01-15 | Sales | | | | sales_q1.xlsx |
| Frank Lee | frank@leads.com | 7200.0 | South | 2024-04-10 | Sales | Top performer | | | sales_q2.xlsx |
| Laura Moss | laura@corp.com | | East | | HR | | 2021-03-15 | 55000 | hr_employees.csv |
| Sam Turner | sam@leads.com | 5400.0 | North | 2024/01/08 | Marketing | | | | marketing_leads.csv |
| Yara Singh | yara@ops.com | 14200.0 | West | 2024-03-18 | Operations | | | | ops_data.xlsx |
| ... | | | | | | | | | |

The **"Merge Summary"** sheet shows per-file row counts and the number of duplicates removed.

## Customizing Column Mapping

To add support for new column name variants, edit the `COLUMN_ALIASES` dictionary in `excel_merger.py`:

```python
COLUMN_ALIASES = {
    "sales rep": "name",      # maps "sales rep" → "name"
    "deal size": "revenue",   # maps "deal size" → "revenue"
    # ...
}
```

## Notes

- `.xls` files (Excel 97-2003) require the `xlrd` package.
- Date normalization is intentionally left as-is to preserve original values — add a date parsing step if needed for your use case.
- For very large files (10,000+ rows), pandas handles them efficiently in memory.
