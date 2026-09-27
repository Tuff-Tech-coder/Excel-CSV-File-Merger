# Excel & CSV File Merger

![Python](https://img.shields.io/badge/Python-3.11+-3776AB?logo=python&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-3.x-150458?logo=pandas&logoColor=white)
![openpyxl](https://img.shields.io/badge/openpyxl-3.x-1D6F42)
![Tests](https://img.shields.io/badge/tests-64%20passing-brightgreen)
![License](https://img.shields.io/badge/License-MIT-green)

A defensive data-normalization pipeline that recursively discovers Excel and CSV exports, aligns their schemas, and vertically concatenates them into a reviewable two-sheet workbook.

The project is designed for real business exports: column aliases, reordered fields, currency stored as text, mixed date formats, empty rows, exact duplicates, nested folders, source lineage, and partially unreadable batches.

## Demonstrated result

The included five-file fixture produces a deterministic audit trail:

- 5 source files discovered and processed
- 29 source rows loaded
- 1 blank row removed
- 1 exact duplicate removed
- 27 normalized rows written
- `Merge Summary` and `Merged Data` worksheets

`Salary` remains separate from `Revenue`, and `Hire Date` remains separate from transaction `Date`; unrelated business concepts are never combined merely because they share a numeric or date type.

## What it handles

- Recursive, case-insensitive discovery of `.xlsx`, `.xls`, and `.csv` files.
- Deterministic source ordering and relative-path lineage in `source_file`.
- Configurable column aliases such as `Sales Amount` → `revenue` and `Territory` → `region`.
- Collision-safe headers, including preservation of a user-supplied `source_file` field.
- Strict money parsing with support for symbols, thousands separators, negatives, and accounting parentheses.
- Typed Excel date cells for recognized date columns while preserving unrecognized source text.
- Leading-zero CSV identifiers such as `00123` without automatic numeric coercion.
- Empty-row cleanup after whitespace trimming.
- Optional exact-row deduplication that excludes only the generated lineage field.
- Formula-injection protection for untrusted headers and cell text.
- Formula cells from `.xlsx` inputs preserved as inert text instead of disappearing or executing.
- Atomic output replacement, rotating logs, structured JSON run summaries, and explicit partial-run status.

## Output workbook

`Merge Summary` is the first worksheet and shows the main run metrics plus a source-by-source audit table. `Merged Data` contains the normalized records with:

- an Excel table and filters;
- a frozen header row;
- typed currency and date formatting;
- bounded, readable column widths; and
- source lineage as the final column.

The output must be outside the input folder. This prevents a source workbook from being overwritten and prevents previous output from being ingested on a later run.

## Setup

Run these commands from the project root.

PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

macOS or Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Usage

Merge the included sample files and save both workbook and JSON audit artifacts:

```bash
python excel_merger.py \
  --input ./sample_input \
  --output ./merged_master.xlsx \
  --json-out ./merge_run.json
```

PowerShell uses the same options on one line:

```powershell
python excel_merger.py --input .\sample_input --output .\merged_master.xlsx --json-out .\merge_run.json
```

Options:

- `--input`: source folder; nested folders are included.
- `--output`: destination `.xlsx` file outside the input folder.
- `--json-out`: optional machine-readable run summary.
- `--no-dedup`: keep exact duplicate rows.
- `--strict`: abort if any discovered input cannot be read.
- `--overwrite`: explicitly replace an existing destination workbook.

Without `--strict`, readable files still produce a workbook when another input fails. The run is marked `partial`, every skipped relative path is listed in the JSON summary, and the CLI exits with status `2` rather than reporting full success.

## How normalization works

1. Discover supported files recursively and reject an output path inside the source tree.
2. Read CSV fields as text to preserve identifiers; read `.xlsx` formulas without executing them.
3. Normalize and de-duplicate headers, trim text, remove blank rows, and type only known money/date fields.
4. Neutralize formula-like text and add relative source lineage.
5. Align columns and concatenate the normalized frames.
6. Optionally remove exact duplicates across all business columns.
7. Write the workbook to a temporary file and atomically replace the destination only after export succeeds.

This is vertical concatenation, not a relational join. Records are stacked after schema alignment; the tool does not match rows by a business key.

## Customizing aliases

`COLUMN_ALIASES` is the main extension point:

```python
COLUMN_ALIASES: dict[str, str] = {
    "full name": "name",
    "contact email": "email",
    "sales amount": "revenue",
    "territory": "region",
    "transaction date": "date",
    "hire date": "hire_date",
    "salary": "salary",
}
```

When two source headers map to the same canonical name, the second receives a deterministic suffix such as `revenue_2`; no column is silently discarded.

## Development

```bash
pip install -r requirements.txt
pip install pytest ruff
pytest -q          # 64 tests
ruff check .
```

The regression suite covers parsing, schema collisions, whitespace-only rows, source lineage, recursive discovery, identifier preservation, formula safety, output-path protection, partial and strict modes, workbook formatting, atomic overwrite behavior, JSON summaries, and the complete five-file demonstration.

## Current scope

- Each `.xlsx` file contributes its active worksheet; legacy `.xls` files contribute their first worksheet.
- The tool recognizes four configured money/date column names (`revenue`, `salary`, `date`, and `hire_date`); additional typed fields should be added deliberately.
- Deduplication removes only exact normalized-row matches. Business-key matching belongs in a separate join/reconciliation workflow.

## License

MIT — see [LICENSE](LICENSE).
