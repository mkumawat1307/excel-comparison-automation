# Excel Comparison Automation

A client-ready Streamlit application for comparing two Excel (.xlsx) files using configurable matching keys and generating a detailed, formatted Excel report.

Designed for business users who need to reconcile two versions of the same dataset — for example, an "old" export vs. a "new" export, or a File A from one system vs. a File B from another — without writing formulas or VBA.

## Problem It Solves

Comparing two Excel files manually is error-prone and time-consuming:

- **VLOOKUP/INDEX-MATCH breaks** when column names differ between files.
- **Composite keys** (multiple columns identifying one record) are difficult to manage with built-in Excel tools.
- **Leading zeros** in identifiers (e.g. `00123`) can be lost if Excel auto-converts text to numbers.
- **Formula injection** risks when cell values begin with `=`, `+`, `-`, or `@`.
- No built-in, auditable report listing every change, missing record, new record, and data-quality issue in one place.

Excel Comparison Automation automates this workflow with configurable key mapping, type-aware comparison (text / numeric / date), and a client-ready six-sheet Excel report.

## Key Features

| Feature | Description |
|---------|-------------|
| Two-file upload | Upload File A and File B (.xlsx only, max 50 MB each). |
| Worksheet selection | Choose one worksheet from each file to compare. |
| Matching keys | Configure one or more column pairs, including differently named columns (e.g. `Site ID` ↔ `Site Code`). |
| Composite keys | Match on multiple key columns combined using collision-safe tuple keys. |
| Auto-suggestion | Exact-name column matches suggested automatically for keys and comparisons. |
| Comparison columns | Select columns to compare for value changes; key columns excluded automatically. |
| Comparison types | Text (default), Numeric (with tolerance), and Date (`pd.to_datetime`, `dayfirst=False`). |
| Six-sheet Excel report | Comparison Overview, Unchanged Records, Records With Changes, Missing From File B, New Records in File B, Records Needing Attention. |
| Formula-injection protection | All text written as explicit OpenPyXL text cells; `=`, `+`, `-`, `@` prefixed values are never interpreted as formulas. |
| Original value preservation | Leading zeros, whitespace, casing, and data types preserved exactly; source workbooks never modified. |
| Excel row references | Every report row references the physical Excel row number from the source file. |
| Duplicate and missing-key handling | Reported as data-quality issues, never silently dropped; unambiguous matching continues non-blocking. |
| Large-file protection | Compressed file size (50 MB), declared decompressed ZIP size (200 MB), and per-worksheet resource limits (250,000 rows, 1,000 columns, 5,000,000 cells) are enforced with clear errors. |

## Technology Stack

| Layer | Technology |
|-------|------------|
| Language | Python >= 3.10 |
| Web UI | Streamlit >= 1.30.0 |
| Data manipulation | pandas >= 2.0.0 |
| Excel reading/writing | openpyxl >= 3.1.0 |
| Testing | pytest >= 8.0.0 |
| Linting/formatting | ruff >= 0.5.0 |

## Screenshots

Screenshots are not yet available. Add them here when you have captured the application UI:

- `docs/screenshots/01-upload.png` — Step 1: File upload
- `docs/screenshots/02-worksheets.png` — Step 2: Worksheet selection
- `docs/screenshots/03-keys.png` — Step 3: Matching keys configuration
- `docs/screenshots/04-comparisons.png` — Step 4: Comparison columns configuration
- `docs/screenshots/05-results.png` — Step 5: Results dashboard and preview tabs
- `docs/screenshots/06-download.png` — Step 6: Report download
- `docs/screenshots/report-overview.png` — Comparison Overview sheet
- `docs/screenshots/report-unchanged.png` — Unchanged Records sheet (Option B layout)

## Installation

```powershell
# Clone the repository
git clone https://github.com/mkumawat1307/excel-comparison-automation.git
cd excel-comparison-automation

# Create and activate a virtual environment
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# Install the package and dev dependencies
pip install -e ".[dev]"
```

## How to Run the Application

```powershell
streamlit run app/streamlit_app.py
```

Then open the URL Streamlit prints (typically `http://localhost:8501`) in your browser.

See [docs/INSTALLATION.md](docs/INSTALLATION.md) for full installation instructions, including prerequisites, virtual environment setup, and troubleshooting.

## Basic Usage Workflow

1. **Upload** File A and File B (`.xlsx`, max 50 MB each).
2. **Select worksheets** — one from each file.
3. **Configure matching keys** — pair the columns that identify a record (auto-suggested exact-name matches available).
4. **Configure comparison columns** — choose which columns to compare for changes, with type (Text / Numeric / Date).
5. **Run comparison** — view the dashboard with summary metrics and five results tabs.
6. **Download** the formatted six-sheet Excel report.

See [docs/USER_GUIDE.md](docs/USER_GUIDE.md) for step-by-step instructions written for non-technical Excel users.

## Generated Excel Report

The report contains six sheets:

| Sheet | Contents |
|-------|----------|
| Comparison Overview | Summary metrics in plain language (records matched, changed, missing, new, duplicates, missing keys, column stats). |
| Unchanged Records | Records with matching keys and equal compared values; interleaved key pairs, File B comparison values, unmapped columns, Excel row references. |
| Records With Changes | Every changed cell: previous vs. updated value, column names, key values, Excel row references. |
| Missing From File B | Complete original File A records (original column names, no prefix) with Excel Row (File A). |
| New Records in File B | Complete original File B records (original column names, no prefix) with Excel Row (File B). |
| Records Needing Attention | Duplicate keys and missing key values, with `Problem Found`, `File`, `Excel Row`, key values, and related column values. |

See [docs/REPORT_GUIDE.md](docs/REPORT_GUIDE.md) for detailed sheet-by-sheet explanations.

## Input Limits and Validation

### File validation

Every uploaded file is validated before processing (see `src/excel_compare/security.py`):

- **Format:** only `.xlsx` workbooks are accepted (case-insensitive extension check).
- **Compressed size:** maximum 50 MB per file (`50 * 1024 * 1024` bytes).
- **Declared decompressed size:** the sum of the ZIP central-directory uncompressed sizes must not exceed 200 MB (`200 * 1024 * 1024` bytes). This bounds ZIP-bomb style expansion before the workbook is opened.
- **Malformed archives:** a file that is not a valid ZIP archive is rejected with a clear error before the readability check.
- **Readability:** the workbook must open via OpenPyXL (`read_only=True, data_only=True`) before further processing.

The decompressed size is read from the archive's central-directory metadata. This is a fast pre-open guard, not a verification of archive integrity or of the actual decompressed bytes.

### Worksheet resource limits

Loading a worksheet enforces the following limits (`src/excel_compare/config.py`). Exceeding any limit raises a clear error naming the limit that was exceeded:

| Limit | Value | What is counted |
|-------|-------|-----------------|
| Yielded rows | 250,000 | Every yielded physical worksheet row: leading blank rows, the header, interior gaps, and empty rows. |
| Actual data rows | 250,000 | Rows containing at least one non-empty value. |
| Columns | 1,000 | Width of the header row and of each yielded data row, including formatting-only cells. |
| Cells | 5,000,000 | Total cells across rows that contain at least one non-empty value. |

How blank and formatting-only content is treated:

- An **empty row** (no non-empty cell values) does not count toward actual data rows or cells, but it does count toward the yielded-row limit. Leading blank rows before the header and the header row itself also count toward the yielded-row limit, which is enforced while the header is being located.
- An **empty cell** inside a data row is preserved as `None`; it contributes to that row's width and cell count.
- A **formatting-only cell** (styled but with no value) contributes to the row width and cell count, but does not by itself make a row count as data.

### Worksheet dimension protection

Excel stores a `<dimension>` element declaring a worksheet's used range, but that metadata cannot always be trusted: it can be **inflated** (claiming a much larger range than exists), **underreported** (declaring fewer rows or columns than actually contain data), or **missing** entirely. Trusting it directly would either reject small valid worksheets or silently drop real data.

Before iterating, the loader calls OpenPyXL's `worksheet.reset_dimensions()` (`src/excel_compare/loader.py`). This clears the cached dimension so `iter_rows()` walks the actual worksheet XML and sizes each row to its real last cell. As a result:

- **Inflated dimensions** no longer cause a small worksheet to be rejected.
- **Underreported dimensions** no longer hide data that exists beyond the declared range.
- **Missing dimensions** load normally.

The resource limits above are then enforced on the real data actually encountered, and any worksheet that exceeds a limit is rejected with a clear error rather than being silently truncated.

## How to Run Tests

```powershell
# Full test suite (303 tests)
python -m pytest -q

# Individual test file
python -m pytest tests/unit/test_writer.py -q

# Lint check
python -m ruff check .

# Format check
python -m ruff format --check .
```

See [docs/TESTING.md](docs/TESTING.md) for details.

## Project Documentation

| Document | Description |
|----------|-------------|
| [docs/PROJECT_OVERVIEW.md](docs/PROJECT_OVERVIEW.md) | Purpose, features, technology stack, scope, capabilities, limitations. |
| [docs/INSTALLATION.md](docs/INSTALLATION.md) | Prerequisites, Python version, virtual environment, dependencies, launch, tests, ruff, troubleshooting. |
| [docs/USER_GUIDE.md](docs/USER_GUIDE.md) | Step-by-step instructions for non-technical Excel users. |
| [docs/COMPARISON_LOGIC.md](docs/COMPARISON_LOGIC.md) | Matching keys, composite keys, text/numeric/date comparison, duplicate/missing-key handling. |
| [docs/REPORT_GUIDE.md](docs/REPORT_GUIDE.md) | All six report sheets, column layouts, metrics, value preservation. |
| [docs/PROJECT_STRUCTURE.md](docs/PROJECT_STRUCTURE.md) | Directory layout and module responsibilities with actual function names. |
| [docs/TESTING.md](docs/TESTING.md) | Test structure, execution commands, regression coverage. |
| [docs/SECURITY.md](docs/SECURITY.md) | Implemented protections and recommended future improvements. |

## Security and Privacy

- All processing is local; no data is transmitted over the network.
- Only `.xlsx` files are accepted, with a 50 MB compressed size limit and a 200 MB declared decompressed ZIP size limit per file.
- Per-worksheet resource limits (250,000 rows, 1,000 columns, 5,000,000 cells) reject oversized input with a clear error instead of truncating it.
- Uploaded files are validated for readability before processing.
- All report text cells use formula-injection protection (`data_type='s'`).
- Source workbooks are opened read-only and never modified.
- No API keys, passwords, or secrets are used or stored.

See [docs/SECURITY.md](docs/SECURITY.md) for implemented protections and recommended future improvements.

## Current Limitations

- Only `.xlsx` files (no `.xls`, `.xlsm`, `.csv`, `.ods`).
- Maximum 50 MB compressed per file and 200 MB declared decompressed ZIP size per file.
- Maximum 250,000 yielded rows, 250,000 actual data rows, 1,000 columns, and 5,000,000 cells per worksheet; oversized worksheets are rejected rather than truncated.
- One worksheet per file per comparison run.
- Exact-name column matching only (no fuzzy matching).
- Comparison configuration is not persisted between browser sessions (Streamlit session state only).
- Numeric and date comparison apply only to explicitly configured pairs; all others use text comparison.
- Date parsing uses `dayfirst=False` only.
- No CLI entry point; runs via `streamlit run app/streamlit_app.py`.
- No license has been selected yet (see below).

## License

No license has been selected yet for this project. The source code is provided as-is without any license grant. If you plan to fork, modify, or redistribute this project, please contact the maintainer to discuss licensing.

## Contributing

This project is being prepared for publication. Contribution guidelines will be added when the project is released. Do not commit or push changes without authorization from the maintainer.