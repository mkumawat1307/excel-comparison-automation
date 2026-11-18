# Report Guide

Excel Comparison Automation generates a formatted `.xlsx` report with six sheets. Sheet names are fixed and never change:

1. Comparison Overview
2. Unchanged Records
3. Records With Changes
4. Missing From File B
5. New Records in File B
6. Records Needing Attention

All text values are written as explicit text cells, so strings beginning with `=`, `+`, `-`, or `@` are never interpreted as Excel formulas. Original values are preserved exactly.

## Sheet 1: Comparison Overview

This sheet summarizes the entire comparison in plain language. It has two columns: **Metric** and **Value**.

| Metric key (internal) | Display label | Meaning |
|-----------------------|---------------|---------|
| `total_rows_file_a` | Records in File A | Total File A rows accounted for (matched + only-in-A + duplicates + missing-key). |
| `total_rows_file_b` | Records in File B | Total File B rows accounted for. |
| `matched_records` | Records with no changes | Matched records where every compared value is equal. |
| `mismatched_records` | Records with changes | Matched records where at least one compared value differs. |
| `matched_no_fields_compared` | Records matched but not compared | Records matched by key but no comparison columns were configured. |
| `only_in_file_a` | Records missing from File B | Records in File A with no matching key in File B. |
| `only_in_file_b` | New records in File B | Records in File B with no matching key in File A. |
| `duplicate_key_records` | Duplicate identification records | Records with duplicate key values (data-quality issues). |
| `missing_key_records` | Records with missing identification values | Records with empty key values (data-quality issues). |
| `duplicate_groups_file_a` | — | Number of duplicate key groups detected in File A. |
| `duplicate_groups_file_b` | — | Number of duplicate key groups detected in File B. |
| `common_columns` | Common columns | Count of column names present in both files. |
| `columns_only_in_file_a` | Columns only in File A | Count of column names present only in File A. |
| `columns_only_in_file_b` | Columns only in File B | Count of column names present only in File B. |

Unknown metric keys fall back to title case (underscores replaced with spaces).

## Sheet 2: Unchanged Records

This sheet lists records whose keys match in both files and whose compared values are all equal.

### Column order (approved layout)

| Section | Columns |
|---------|---------|
| 1. Interleaved key column pairs | `File A – {key1}`, `File B – {key1}`, `File A – {key2}`, `File B – {key2}`, … (one pair per configured key) |
| 2. File B comparison columns | `File B – {compare_col}` for each selected comparison column, in configured order (File B name only; File A comparison columns are omitted since values are identical) |
| 3. Status and row references | `Record Status`, `Excel Row (File A)`, `Excel Row (File B)` |

Only the configured key columns and selected comparison columns are shown. Unmapped columns (present in a source file but not selected as a key or comparison column) are **not** displayed on this sheet.

### What each column means

- **File A – {key} / File B – {key}**: The key value from each file for this record. Values are interleaved by key-pair order so you can confirm both sides match.
- **File B – {compare_col}**: The comparison value from File B. Because this is the Unchanged Records sheet, this equals the File A value (not shown separately to avoid duplication).
- **Record Status**: Plain-language status label (for unchanged records: "No changes").
- **Excel Row (File A) / Excel Row (File B)**: Physical row number in the source worksheet (the actual worksheet row, including any leading blank rows before the header).

## Sheet 3: Records With Changes

This sheet lists every cell that differs between matched records. One row per changed cell.

### Columns

| Column | Meaning |
|--------|---------|
| `File A – {key1}`, `File B – {key1}`, … | Interleaved key values identifying the record. |
| `Column Name (File A)` | The File A column name for the changed cell. |
| `Previous Value` | The original value from File A. |
| `Column Name (File B)` | The File B column name for the changed cell. |
| `Updated Value` | The new value from File B. |
| `Excel Row (File A)` | Physical row number in File A. |
| `Excel Row (File B)` | Physical row number in File B. |

If a record has multiple changed columns, it appears as multiple consecutive rows (one per changed cell), with the same key values repeated.

## Sheet 4: Missing From File B

This sheet lists records that exist in File A but have no matching key in File B.

### Layout (complete original record)

| Section | Columns |
|---------|---------|
| 1. All File A source columns | Every column from the File A worksheet, using **original column names** with no prefix, in original workbook order. |
| 2. Row reference | `Excel Row (File A)` — physical row number in File A. |

### What is NOT included

- No File B columns (there is no matching record in File B).
- No `Record Status` column.
- No artificial `Record ID` or `Row Key` column.
- No `File A –` prefix on column names.

Each source column appears exactly once — key columns are not duplicated.

## Sheet 5: New Records in File B

This sheet lists records that exist in File B but have no matching key in File A.

### Layout (complete original record)

| Section | Columns |
|---------|---------|
| 1. All File B source columns | Every column from the File B worksheet, using **original column names** with no prefix, in original workbook order. |
| 2. Row reference | `Excel Row (File B)` — physical row number in File B. |

### What is NOT included

- No File A columns (there is no matching record in File A).
- No `Record Status` column.
- No artificial `Record ID` or `Row Key` column.
- No `File B –` prefix on column names.

Each source column appears exactly once — key columns are not duplicated.

## Sheet 6: Records Needing Attention

This sheet lists data-quality problems found during matching: duplicate keys and missing key values.

### Columns

| Column | Meaning |
|--------|---------|
| `Problem Found` | Issue type: `DUPLICATE_KEY` or `MISSING_KEY`. |
| `File` | Source file: `A` or `B`. |
| `Excel Row` | Physical row number in the source worksheet. |
| `File A – {key1}`, `File B – {key1}`, … | Interleaved key values for this issue (one side may be blank if the issue is only on one file). |
| `Related Column Values` | All column values for the record as `Column = 'value'` pairs, joined by ` | `, sorted alphabetically by column name. |

### Issue types

- **DUPLICATE_KEY**: The same composite key appears more than once in a file. These records are not matched because the application cannot determine which occurrence corresponds to which. Review manually.
- **MISSING_KEY**: A record has an empty or missing value in a key column. These records are excluded from matching.

## Value Preservation

- Original column names are preserved exactly (no title-casing, no acronym expansion, no whitespace trimming).
- Original cell values are preserved exactly (leading zeros, spaces, casing).
- Data types from the source workbook are preserved as text cells so values like `00123` or `=1+1` are not altered by Excel.

## Source Code References

| Behavior | Module | Function |
|----------|--------|----------|
| Six fixed sheet names | `config.py` | `REPORT_SHEETS` |
| Plain-language summary labels | `reporter.py` | `SUMMARY_LABELS`, `summary_label()` |
| Per-file column headings (en dash) | `reporter.py` | `file_a_column_header()`, `file_b_column_header()`, `FILE_COLUMN_SEPARATOR` |
| Unchanged Records sheet (keys + File B comparisons only) | `writer.py` | `_write_matching_records_sheet()` |
| Records With Changes sheet | `writer.py` | `_write_mismatched_records_sheet()` |
| Missing / New sheets (complete records) | `writer.py` | `_write_only_in_sheet()` |
| Records Needing Attention sheet | `writer.py` | `_write_data_quality_sheet()` |
| Formula-injection protection | `writer.py` | `_write_text_cell()` (`data_type='s'`) |
| Report orchestration | `writer.py` | `write_report()` |