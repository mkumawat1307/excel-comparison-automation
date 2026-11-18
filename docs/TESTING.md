# Testing

Excel Comparison Automation uses pytest for unit testing and ruff for linting and formatting. All tests live in `tests/unit/` with shared fixtures in `tests/conftest.py`.

## Test Structure

| Test file | Module under test | What it covers |
|-----------|-------------------|----------------|
| `tests/unit/test_loader.py` | `loader.py` | Valid file loading, file-like objects, sheet listing, missing worksheet, empty worksheet, required columns, column selection, value preservation (leading zeros, whitespace), invalid/nonexistent files, repeated reads. |
| `tests/unit/test_normalizer.py` | `normalizer.py` | Missing-value detection (None, NaN, pd.NA, pd.NaT, empty string), whitespace normalization, case sensitivity, empty-string preservation, numeric/datetime type preservation, immutability of input DataFrame. |
| `tests/unit/test_matcher.py` | `matcher.py` | Single and composite key matching, collision-safe tuple keys (values with `|`), missing key column errors, empty DataFrames, missing/None key values preserved, duplicate keys (non-blocking), unsupported `all` policy, duplicate detection indices, original index preservation, deterministic schema diff ordering, duplicate source record values, row-accounting invariant, renamed key pairs, mapping validation errors, auto-pair exact-name matching. |
| `tests/unit/test_comparer.py` | `comparer.py` | Exact match, numeric within/outside tolerance, tolerance boundary, leading-zero identifiers not auto-converted to numeric, explicit date parsing, date mismatch, unparseable dates, text comparison default, missing vs missing/present/empty, case sensitivity, whitespace normalization, original value preservation in CellDiff, row immutability, renamed compare pairs, numeric/date pairs with different names, empty compare pairs. |
| `tests/unit/test_reporter.py` | `reporter.py` | Basic report assembly, mismatched records, duplicate vs missing-key distinction, schema diff preservation, only-in-A/only-in-B, summary row accounting, renamed key/compare pairs, empty compare pairs → matched-not-compared, unmapped columns not counted as mismatches, per-file key value preservation, summary labels. |
| `tests/unit/test_writer.py` | `writer.py` | Six sheet names unchanged, data quality sheet headings (`Problem Found`, `File`, `Excel Row`, `Related Column Values`, no `Record ID`), original value preservation, formula injection (`=`, `+`, `-`, `@` stored as literal text with `data_type='s'`), key values populated on Unchanged Records, composite keys, renamed keys, backward compatibility without key_values, Records With Changes headings, formula injection on mismatched sheet, Missing / New sheets use original column names (no prefix, no `Record Status`), Comparison Overview plain-language labels, per-file key columns regression, complete-record regression, column heading regression (en dash, original names preserved), Excel row heading consistency, formula injection still protected. |
| `tests/unit/test_models.py` | `models.py` | DiffStatus enum values and plain-language labels, CellDiff/RowDiff defaults, ComparisonResult summary default, DataQualityIssue key_values defaults. |
| `tests/unit/test_security.py` | `security.py` | Nonexistent file, directory path, unsupported extension (`.csv`), empty file, unreadable workbook, validation workbook closed, file handle released for deletion. |
| `tests/unit/test_streamlit_app.py` | `streamlit_app.py` | App loads without errors, file upload widgets render, session state initialized, navigation steps, no crashes on rerun, Step 3/4 column-pair iteration, auto-suggestion checkboxes, add/remove pairs, duplicate prevention, type changes, navigation persistence, dashboard four signals (completed, differences, not-compared separate, data quality), end-to-end comparison with widget-selected keys, heading and complete-record regression tests for Streamlit, temporary data cleanup (Finish & Clear Data / Start Over confirmation, all app-owned state reset, uploader nonce rotation, old files not re-ingested, session isolation), Step 1 upload caption and `.streamlit/config.toml` `server.maxUploadSize = 50`. |

Shared fixture (`tests/conftest.py`):
- `tmp_dir` — temporary directory path.
- `make_excel_file(sheets, filename)` — creates an `.xlsx` workbook from a dict of sheet_name → list-of-dicts rows.

## How to Execute Tests

All commands are run from the project root directory with the virtual environment activated (see [INSTALLATION.md](INSTALLATION.md)).

### Run the full test suite

```powershell
python -m pytest -q
```

Expected result: all tests pass. As of the current release: **331 passed, 0 failed**.

### Run an individual test file

```powershell
python -m pytest tests/unit/test_writer.py -q
```

Replace `test_writer.py` with any other test file name.

### Run tests matching a keyword

```powershell
python -m pytest -q -k "duplicate"
```

### Run a single test function

```powershell
python -m pytest tests/unit/test_writer.py::test_write_report_creates_six_sheets -q
```

## Ruff Checks

Ruff is configured in `pyproject.toml` (`line-length = 100`, `target-version = "py310"`, selects `E`, `F`, `I`, `B`, `C4`, `PLC`, `TRY`, ignores `E501`).

### Lint check

```powershell
python -m ruff check .
```

Expected: "All checks passed!"

### Format check

```powershell
python -m ruff format --check .
```

Expected: "24 files already formatted" (or the current file count).

### Apply formatting (if files are not formatted)

```powershell
python -m ruff format .
```

## Regression Scenarios Covered

The following behaviors are protected by regression tests:

- **Six fixed sheet names** — `'Comparison Overview'`, `'Unchanged Records'`, `'Records With Changes'`, `'Missing From File B'`, `'New Records in File B'`, `'Records Needing Attention'`.
- **No combined `Record ID` / `Row Key` column** — per-file key columns are used instead.
- **En-dash column headings** — `File A – {col}`, `File B – {col}` with original column names preserved exactly (spaces, underscores, acronyms unchanged).
- **Plain-language summary labels** — e.g. `matched_records` → "Records with no changes".
- **Option B Unchanged Records layout** — interleaved key pairs + File B comparison columns only + unmapped columns + status + Excel row refs.
- **Complete original records on Missing / New sheets** — original column names, no prefix, no `Record Status`, Excel row reference only.
- **No column loss or duplication** — every source column classified as key, comparison, or unmapped appears exactly once.
- **Formula-injection protection** — `=`, `+`, `-`, `@` prefixed strings stored as literal text cells (`data_type='s'`) on all sheets.
- **Original values preserved** — leading zeros (`00123`), whitespace, casing.
- **Duplicate key handling** — non-blocking; unambiguous records still match; duplicates reported as data-quality issues.
- **Missing key handling** — excluded from matching, reported separately.
- **Row accounting invariant** — every source row accounted for exactly once across categories.
- **Numeric tolerance** — boundary conditions at exactly `1e-9` and just outside.
- **Date parsing** — `dayfirst=False`; unparseable dates are mismatches.
- **Text comparison default** — `100` vs `100.0` are different unless configured as Numeric.
- **Streamlit widget persistence** — add/remove key and compare pairs survive reruns and navigation.

## What Tests Do NOT Cover

The following are explicitly NOT covered by the current test suite. Do not claim they are tested:

- UI rendering of individual DataFrame cells inside `st.dataframe` tabs (AppTest cannot inspect rendered table cell contents).
- End-to-end file upload through the browser (tests simulate session state and button clicks, not actual file uploads).
- Multi-sheet workbooks compared in a single run.
- File sizes near the 50 MB limit.
- Performance or memory usage for large workbooks.
- Password-protected or macro-enabled workbars (`.xlsm`).
- Deployment to cloud platforms or CI/CD pipelines.