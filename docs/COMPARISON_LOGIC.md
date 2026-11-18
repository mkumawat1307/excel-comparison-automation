# Comparison Logic

This document explains how the comparison engine actually works, based on the source code in `src/excel_compare/`.

## Pipeline Overview

The comparison runs in three stages:

1. **Loading** (`loader.py`): Worksheets are read into DataFrames. Original column names, values, and data types are preserved. Physical Excel row numbers are tracked separately (row 1 = header, first data row = row 2).
2. **Matching** (`matcher.py`): Rows are classified as matched, only-in-A, only-in-B, duplicate, or missing-key based on configurable key column pairs.
3. **Comparison** (`comparer.py`): Matched rows are compared cell-by-cell across configured comparison column pairs.
4. **Report assembly** (`reporter.py`): Results are assembled into a `ComparisonResult` with plain-language status labels and summary metrics.
5. **Report writing** (`writer.py`): The result is written to a formatted six-sheet `.xlsx` file.

## Matching Keys

### Configuration

Keys are defined as `ColumnPair` objects, each pairing a File A column with a File B column. Columns can have different names on each side (for example, `Site ID` ↔ `Site Code`).

At least one key pair is required. The UI auto-suggests exact-name matches (columns with identical names in both files), but you may add or remove pairs manually.

### Composite Keys

Multiple key pairs can be configured. The application builds a composite key tuple for each row, one value per key column, in pair order. Composite keys use Python tuples internally — never delimiter-joined strings — so values containing delimiter characters (for example, `A|B`) cannot cause false collisions.

### Independent Column Mapping

File A and File B columns are mapped independently. A File A column cannot appear more than once across key and compare mappings, and the same applies to File B. A column cannot be both a matching key and a comparison column.

### Exact-Name Automatic Suggestions

The `discover_auto_pairs()` function returns `ColumnPair(col, col)` for every column name present in both files, preserving File A order. This is exact-name matching only — no fuzzy matching, no case-insensitive matching, no whitespace-insensitive matching.

## Text Comparison

### Default Behavior

All comparison column pairs that are not explicitly configured as Numeric or Date use text comparison.

### Normalization Applied for Comparison Only

Normalization is applied to values **during comparison only**. The original DataFrame and original cell values are never modified.

| Setting | Default | Effect on comparison |
|---------|---------|----------------------|
| Case sensitivity | `False` | Text is lowercased before comparing (`ALICE` = `alice`). |
| Whitespace normalization | `True` | Leading/trailing whitespace is stripped (` Alice ` = `Alice`). |
| Empty string treated as missing | `False` | Empty string `''` is a real value, not missing (`''` ≠ `None`). |

Original values (with original casing and whitespace) are preserved in `CellDiff` for reporting.

### Text Comparison Rules

- Both values missing (None/NaN/pd.NA) → equal.
- One value missing, one present → mismatch.
- Both values empty strings → equal.
- Otherwise, normalized text values are compared for equality.

## Numeric Comparison

Numeric comparison is performed **only** on column pairs explicitly listed in `numeric_pairs`. String values like `00123` are never auto-converted to numbers.

### Rules

- Both values parse as numbers → compare with `abs(n_a - n_b) <= tolerance`. Default tolerance is `1e-9`.
- One or both values cannot be parsed as numeric → mismatch (original values reported).
- Missing values follow the standard missing-value rules (missing vs missing = equal; missing vs present = mismatch).

Original string values are preserved in the report (for example, `100.000000001` and `100.0` are reported exactly as entered, even though they compare as equal within tolerance).

## Date Comparison

Date comparison is performed **only** on column pairs explicitly listed in `date_pairs`.

### Rules

- Values are parsed with `pd.to_datetime(stripped, dayfirst=False)`.
- Both values parse successfully → compare for equality.
- One or both values fail to parse → mismatch.
- Missing values follow the standard missing-value rules.

Original string values are preserved in the report.

## Unchanged Records

A matched record (same key in both files, appearing exactly once in each) is "unchanged" when:

- At least one comparison column is configured, **and**
- Every compared column pair is equal under the configured rules.

Unchanged records appear on the **Unchanged Records** sheet and are counted as `matched_records` in the summary.

## Changed Records

A matched record is "changed" when at least one compared column pair differs. Each differing cell appears as a row on the **Records With Changes** sheet, showing:

- The key values (interleaved File A / File B).
- Column Name (File A), Previous Value (from File A).
- Column Name (File B), Updated Value (from File B).
- Excel Row (File A), Excel Row (File B).

## Missing Records

### Missing From File B

A record in File A whose key does not appear in File B (and is not duplicated or missing-key) is classified as "only in A" and appears on the **Missing From File B** sheet as a complete original File A record with original column names and Excel Row (File A).

### New Records in File B

A record in File B whose key does not appear in File A (and is not duplicated or missing-key) is classified as "only in B" and appears on the **New Records in File B** sheet as a complete original File B record with original column names and Excel Row (File B).

## Duplicate Keys

When the same composite key value appears more than once in File A, File B, or both:

- None of the records in the ambiguous group are automatically matched.
- Every affected source record is preserved with original values.
- Duplicate groups are reported in **Records Needing Attention** with `issue_type = DUPLICATE_KEY`.
- Unambiguous matching continues independently for other keys (non-blocking behavior).

Example: File A has two rows with `id = 1`, File B has one row with `id = 1`. None of the three records is matched. All three are reported as duplicates.

## Missing Keys

When any key column in a row is empty, None, NaN, or whitespace-only:

- The record is excluded from matching.
- It is reported as a `MISSING_KEY` data-quality issue in **Records Needing Attention**.
- The record's original values and Excel row reference are preserved.

Missing-key records never appear in matched, only-in-A, or only-in-B categories.

## Matched Records with No Fields Selected for Comparison

If no comparison columns are configured:

- Matched pairs are classified as `MATCHED_NO_FIELDS_COMPARED` (display label: "Matched, not compared").
- They are **not** counted as unchanged records.
- They appear as a separate informational signal on the dashboard.
- The `compare_rows()` function returns an empty list of cell diffs (zero comparisons performed).

## Row Accounting

Every source row is accounted for exactly once across categories. For each file:

```
total_rows = matched_pairs + only_in_unique + duplicate_source_records + missing_key_records
```

This invariant is tested and enforced in the matcher tests.

## Source Code References

| Behavior | Module | Function / Class |
|----------|--------|------------------|
| Column pair auto-discovery (exact name) | `matcher.py` | `discover_auto_pairs()` |
| Column mapping validation | `matcher.py` | `validate_column_mapping()` |
| Key tuple building and normalization | `matcher.py` | `_build_key_tuples()`, `_key_is_missing()` |
| Duplicate detection | `matcher.py` | `detect_duplicates()` |
| Row matching and classification | `matcher.py` | `match_on_keys()` |
| Schema diff computation | `matcher.py` | `build_schema_diff()` |
| Cell-by-cell comparison | `comparer.py` | `compare_rows()` |
| Numeric parsing and tolerance | `comparer.py` | `_try_parse_numeric()`, tolerance from config |
| Date parsing | `comparer.py` | `_try_parse_date()`, `dayfirst=False` |
| Text normalization | `comparer.py` | `_normalize_text()` |
| Missing value detection | `comparer.py`, `normalizer.py` | `_is_missing()`, `canonical_missing()` |
| Report assembly and status labels | `reporter.py` | `build_report()`, `DiffStatus`, `SUMMARY_LABELS` |
| Report writing | `writer.py` | `write_report()`, sheet writer functions |
| Original value preservation on load | `loader.py` | `load_sheet_with_metadata()` |