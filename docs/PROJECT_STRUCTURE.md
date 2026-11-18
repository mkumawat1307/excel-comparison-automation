# Project Structure

```
excel-comparison-automation/
├── pyproject.toml          # Project metadata, dependencies, ruff config
├── README.md               # Project landing page
├── docs/                   # This documentation directory
├── app/
│   └── streamlit_app.py    # Streamlit UI (six-step wizard)
├── src/
│   └── excel_compare/      # Core comparison library
│       ├── __init__.py     # Package version (__version__ = '0.1.0')
│       ├── config.py       # Defaults and constants
│       ├── exceptions.py   # Custom exception classes
│       ├── models.py       # Typed data models
│       ├── loader.py       # Excel file loading
│       ├── normalizer.py   # Value normalization
│       ├── matcher.py      # Composite-key matching
│       ├── comparer.py     # Per-cell comparison
│       ├── reporter.py     # Report assembly
│       ├── writer.py       # Excel report writing
│       └── security.py     # File upload validation
└── tests/
    ├── conftest.py         # Shared pytest fixtures
    └── unit/
        ├── test_loader.py
        ├── test_normalizer.py
        ├── test_matcher.py
        ├── test_comparer.py
        ├── test_reporter.py
        ├── test_writer.py
        ├── test_models.py
        ├── test_security.py
        └── test_streamlit_app.py
```

## Module Responsibilities

### `src/excel_compare/config.py`

Application defaults and constants. No external configuration files are read; all settings are defined in Python.

Key exports:
- `DEFAULT_KEYS`, `DEFAULT_COMPARE_COLUMNS`, `DEFAULT_NUMERIC_COLUMNS`, `DEFAULT_DATE_COLUMNS` — legacy defaults (empty lists / `None`).
- `DEFAULT_KEY_PAIRS`, `DEFAULT_COMPARE_PAIRS`, `DEFAULT_NUMERIC_PAIRS`, `DEFAULT_DATE_PAIRS` — pair-based defaults (empty lists).
- `DEFAULT_CASE_SENSITIVE = False`, `DEFAULT_NORMALIZE_WHITESPACE = True`, `DEFAULT_NUMERIC_TOLERANCE = 1e-9`, `DEFAULT_DUPLICATE_POLICY = 'error'`, `VALID_DUPLICATE_POLICIES = ('error',)`.
- `DEFAULT_MAX_FILE_SIZE_MB = 50`, `ALLOWED_EXTENSIONS = ('.xlsx',)`, `MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024`.
- `REPORT_SHEETS` — tuple of the six fixed sheet names:
  `'Comparison Overview'`, `'Unchanged Records'`, `'Records With Changes'`, `'Missing From File B'`, `'New Records in File B'`, `'Records Needing Attention'`.

### `src/excel_compare/exceptions.py`

Custom exception hierarchy:

- `ExcelCompareError` — base exception.
- `LoadError` — file cannot be loaded or read.
- `ConfigError` — invalid configuration.
- `SecurityError` — file fails security validation.
- `DuplicateKeyError` — duplicate keys detected and policy is unsupported.
- `InvalidKeySelectionError` — key column missing from one or both files.
- `MappingValidationError` — column mapping configuration is invalid.

### `src/excel_compare/models.py`

Typed data models (frozen dataclasses):

- `DiffStatus` — enum: `MATCHED`, `MISMATCHED`, `ONLY_IN_A`, `ONLY_IN_B`, `DUPLICATE`, `MISSING_KEY`, `MATCHED_NO_FIELDS_COMPARED`. Each has a `.label` property returning a plain-language display string.
- `ColumnPair(file_a_column, file_b_column)` — maps one File A column to one File B column.
- `CellDiff` — result of comparing one cell pair: `column_a`, `column_b`, `value_a`, `value_b`, `are_equal`, row/Excel row references, `row_key`, `sheet`.
- `RowDiff` — one compared record: `row_key`, `status`, `cell_diffs`, `index_a`, `index_b`, `excel_row_a`, `excel_row_b`, `key_values_a`, `key_values_b`.
- `DuplicateRecord`, `ComparisonConfig`, `LoadedSheet(df, sheet_name, file_source, excel_row_mapping)`, `SchemaDiff`, `DuplicateGroup`, `DuplicateSourceRecord`, `MissingKeyRecord`, `DataQualityIssue`, `MatchResult`, `ComparisonResult`.

### `src/excel_compare/loader.py`

Excel file loading. Reads workbooks into DataFrames while preserving original column names, values, and data types. Physical Excel row numbers are tracked separately and never mutate the DataFrame.

Key functions:
- `list_sheets(file_source)` — returns worksheet names in a workbook.
- `load_sheet(file_source, sheet_name, columns=None, required_columns=None)` — loads a worksheet into a DataFrame.
- `load_sheet_with_metadata(file_source, sheet_name, columns=None, required_columns=None, track_excel_row=True)` — loads a worksheet into a `LoadedSheet` with `excel_row_mapping` (DataFrame index → physical Excel row number).
- `_open_workbook(file_source)` — opens with openpyxl in read-only, data_only mode; file-like objects are buffered into BytesIO for reuse.

All values are converted to strings on load (preserving leading zeros in identifiers). Truly blank trailing rows are skipped.

### `src/excel_compare/normalizer.py`

Value normalization. Non-mutating: returns a copy, never modifies the input DataFrame. Only object/string columns are transformed; numeric and datetime columns are left untouched.

Key functions:
- `canonical_missing(value)` — returns `True` for None, NaN, pd.NA, pd.NaT. Empty strings are NOT treated as missing. Non-scalars return `False`.
- `normalize_dataframe(df, case_sensitive=False, normalize_whitespace=True, treat_empty_as_missing=False)` — returns a normalized copy. Strips whitespace and/or lowercases string columns based on flags.

### `src/excel_compare/matcher.py`

Composite-key matching. Joins two DataFrames on configurable key column pairs, classifies rows, and detects duplicate and missing keys. Duplicate handling is non-blocking: unambiguous records continue to match normally.

Key functions:
- `discover_auto_pairs(columns_a, columns_b)` — exact-name ColumnPair suggestions (File A order preserved).
- `validate_column_mapping(key_pairs, compare_pairs, numeric_pairs, date_pairs, columns_a, columns_b)` — enforces the mapping contract (existence, no duplicates, no key/compare overlap, numeric/date membership).
- `_config_key_columns()`, `_config_compare_columns()`, `_config_numeric_pairs()`, `_config_date_pairs()` — resolve config fields, handling legacy same-name fields.
- `_build_key_tuples(df, columns, ...)` — builds normalized composite key tuples for matching identity only.
- `_key_is_missing(key_tuple)` — True if any key element is missing or empty.
- `detect_duplicates(key_tuples)` — returns mapping of duplicate key → row indices (collision-safe tuples).
- `match_on_keys(df_a, df_b, key_pairs=None, keys=None, config=None, excel_row_mapping_a=None, excel_row_mapping_b=None)` — classifies rows into matched / only-in-A / only-in-B / duplicate / missing-key. Returns a `MatchResult`.

### `src/excel_compare/comparer.py`

Per-cell comparison of matched rows. Type resolution rules:

- Numeric comparison ONLY on explicitly configured `numeric_pairs`.
- Date comparison ONLY on explicitly configured `date_pairs` (using `pd.to_datetime` with `dayfirst=False`).
- All other pairs use text comparison.

Key functions:
- `compare_rows(row_a, row_b, column_pairs, config, sheet='', row_key='', excel_row_mapping_a=None, excel_row_mapping_b=None)` — compares two rows cell-by-cell. Returns one `CellDiff` per column pair. Original values are preserved; neither row is modified.
- Helper functions: `_is_missing()`, `_empty_string()`, `_normalize_text()`, `_try_parse_date()`, `_try_parse_numeric()`.

### `src/excel_compare/reporter.py`

Report assembly. Builds the `ComparisonResult` from matcher and comparer output.

Key functions:
- `file_a_column_header(column_name)` — returns `File A – {column_name}` (en dash separator).
- `file_b_column_header(column_name)` — returns `File B – {column_name}` (en dash separator).
- `FILE_COLUMN_SEPARATOR = ' – '` — the en-dash separator constant.
- `summary_label(key)` — returns plain-language label for a summary metric key.
- `SUMMARY_LABELS` — dict mapping internal metric keys to display labels.
- `build_report(match_result, cell_diffs_by_pair, config, excel_row_mapping_a=None, excel_row_mapping_b=None)` — assembles `ComparisonResult`: classifies matched pairs as matched / mismatched / matched-no-fields-compared; builds only-in-A / only-in-B lists; converts duplicates and missing keys into `DataQualityIssue` objects; computes summary metrics.

### `src/excel_compare/writer.py`

Excel report writing. Writes the `ComparisonResult` to a formatted `.xlsx` file with six sheets.

Key functions:
- `_write_text_cell(worksheet, row_num, col_num, value)` — writes a value as an explicit OpenPyXL text cell (`data_type='s'`). Formula-injection protection: strings starting with `=`, `+`, `-`, `@` are never executable formulas. Original string values preserved exactly; no apostrophe or tab prefix added.
- `_style_header_row(worksheet, row_num, column_count)` — applies blue header styling.
- `_autofit_columns(worksheet, max_width=60)` — best-fit column widths.
- `_write_summary_sheet(workbook, result)` — Comparison Overview sheet.
- `_write_matching_records_sheet(workbook, result, config, df_a=None, df_b=None)` — Unchanged Records sheet (Option B layout: interleaved keys + File B comparison columns + unmapped columns).
- `_write_mismatched_records_sheet(workbook, result, config)` — Records With Changes sheet.
- `_write_only_in_sheet(workbook, sheet_name, rows, side, config, df=None)` — Missing From File B / New Records in File B sheets (complete original record, original column names, no prefix).
- `_write_data_quality_sheet(workbook, result, config)` — Records Needing Attention sheet.
- `write_report(result, output_path, config, df_a=None, df_b=None)` — writes all six sheets to the output `.xlsx` file. Backward compatible: `df_a`/`df_b` are optional.

### `src/excel_compare/security.py`

Security validation for uploaded Excel files.

Key functions:
- `validate_uploaded_file(file_path)` — validates existence, not-a-directory, `.xlsx` extension (case-insensitive), non-empty, size ≤ 50 MB, readability via openpyxl. Returns resolved `Path` or raises `SecurityError`.
- `safe_copy_to_temp(source_path, temp_dir)` — copies a validated file to an isolated temp location without modifying the original.

### `app/streamlit_app.py`

Streamlit entry point. Six-step wizard UI for non-technical users.

Key functions:
- `_init_session_state()` — initializes all session-state keys.
- `_validate_upload(uploaded_file)` — validates uploaded file (size, extension, readability).
- `_load_worksheet(uploaded_file, sheet_name)` — loads a worksheet from session-state bytes.
- `_build_comparison_config()` — builds `ComparisonConfig` from session-state mappings.
- `_run_comparison()` — runs the full pipeline: match, compare, build report, write to temp file, store bytes in session state.
- `_render_step1_upload()` through `_render_step6_download()` — renders each wizard step.
- `_render_results_preview(result)` — renders the five results tabs (Unchanged Records, Records With Changes, Missing From File B, New Records in File B, Records Needing Attention).
- `main()` — page config, step indicator, step navigation.

## Tests

| Test file | Covers |
|-----------|--------|
| `tests/unit/test_loader.py` | Excel loading, sheet listing, column selection, required columns, empty worksheets, value preservation, file-like objects. |
| `tests/unit/test_normalizer.py` | Missing detection, whitespace/case normalization, empty-string handling, type preservation, immutability. |
| `tests/unit/test_matcher.py` | Single/composite key matching, collision-safe tuples, missing keys, duplicates (non-blocking), invalid keys, mapping validation, auto-pairs, row accounting, schema diff. |
| `tests/unit/test_comparer.py` | Text/numeric/date comparison, tolerance, missing vs present, empty strings, case sensitivity, whitespace, renamed pairs, original value preservation. |
| `tests/unit/test_reporter.py` | Report assembly, status classification, summary metrics, schema diff preservation, plain-language labels, per-file key values. |
| `tests/unit/test_writer.py` | Six sheet names, heading conventions, Option B layout, complete original records, formula injection, key value preservation, backward compatibility. |
| `tests/unit/test_models.py` | DiffStatus values and labels, model defaults, immutability. |
| `tests/unit/test_security.py` | File existence, directory rejection, unsupported extension, empty file, unreadable workbook. |
| `tests/unit/test_streamlit_app.py` | App loading, widget rendering, session state, key/compare pair add/remove, navigation persistence, dashboard signals, end-to-end comparison. |

Total: 234 tests (as of the current release), all passing.

## Configuration File

`pyproject.toml` defines:
- Project name: `excel-comparison-automation`, version `0.1.0`, description `Excel Comparison Automation for business users`.
- Python requirement: `>=3.10`.
- Runtime dependencies: `pandas>=2.0.0`, `openpyxl>=3.1.0`, `streamlit>=1.30.0`.
- Dev dependencies (extra `dev`): `pytest>=8.0.0`, `ruff>=0.5.0`.
- Ruff config: `line-length = 100`, `target-version = "py310"`, selects `E`, `F`, `I`, `B`, `C4`, `PLC`, `TRY`, ignores `E501`.