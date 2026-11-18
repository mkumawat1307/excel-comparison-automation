# Project Overview

## Project Purpose

Excel Comparison Automation is a Streamlit-based desktop application that compares two Excel (.xlsx) workbooks using configurable matching keys and generates a detailed, formatted Excel report highlighting differences.

The application helps business users answer questions such as:
- Which records exist in File A but are missing from File B?
- Which records are new in File B?
- Which records have changed values, and what were the previous vs. updated values?
- Which records have data-quality problems such as duplicate keys or missing identification values?

## Problem This Application Solves

Business teams regularly need to reconcile two versions of the same dataset (for example, an "old" export vs. a "new" export, or a File A from one system vs. a File B from another). Doing this manually in Excel is error-prone and time-consuming:

- VLOOKUP/INDEX-MATCH formulas break when column names differ between files.
- Composite keys (multiple columns identifying one record) are difficult to manage with built-in Excel tools.
- Formula injection risks arise when cell values begin with `=`, `+`, `-`, or `@`.
- Leading zeros in identifiers (for example, `00123`) can be lost if Excel auto-converts text to numbers.
- There is no built-in, auditable report that lists every change, missing record, new record, and data-quality issue in one place.

This application automates that workflow with configurable key mapping, type-aware comparison, and a client-ready Excel report.

## Main Features

| Feature | Description |
|---------|-------------|
| Two-file upload | Upload File A and File B (.xlsx only, max 50 MB each). |
| Worksheet selection | Choose which worksheet to compare from each file. |
| Matching keys | Configure one or more column pairs (including differently named columns) that uniquely identify a record. |
| Composite keys | Match on multiple key columns combined; uses collision-safe tuple keys, not delimiter-joined strings. |
| Auto-suggestion | Exact-name column matches are suggested automatically for keys and comparison columns. |
| Comparison columns | Select which columns to compare for value changes; key columns are excluded automatically. |
| Comparison types | Text (default), Numeric (with configurable tolerance), and Date (using `pd.to_datetime`, `dayfirst=False`). |
| Six-sheet Excel report | Comparison Overview, Unchanged Records, Records With Changes, Missing From File B, New Records in File B, Records Needing Attention. |
| Formula-injection protection | All text values written as explicit OpenPyXL text cells (`data_type='s'`); strings starting with `=`, `+`, `-`, `@` are never interpreted as formulas. |
| Original value preservation | Leading zeros, whitespace, casing, and data types are preserved exactly; original workbooks are never modified. |
| Excel row references | Every report row references the physical Excel row number from the source file. |
| Duplicate and missing-key handling | Duplicate keys and missing key values are reported as data-quality issues, never silently dropped. |
| Non-blocking duplicates | Unambiguous records continue to match normally even when other keys are duplicated. |

## Technology Stack

| Layer | Technology | Version constraint |
|-------|------------|--------------------|
| Language | Python | >= 3.10 |
| Web UI | Streamlit | >= 1.30.0 |
| Data manipulation | pandas | >= 2.0.0 |
| Excel reading/writing | openpyxl | >= 3.1.0 |
| Testing | pytest | >= 8.0.0 (dev) |
| Linting/formatting | ruff | >= 0.5.0 (dev) |

Project metadata: name `excel-comparison-automation`, version `0.1.0`.

## Project Scope

### In Scope

- Loading `.xlsx` workbooks and selecting one worksheet per file.
- Configuring matching key pairs (same-name or differently named columns).
- Configuring comparison column pairs with type (Text / Numeric / Date).
- Running a row-by-row comparison and classifying each record.
- Generating a formatted six-sheet `.xlsx` report.
- Downloading the report from the Streamlit UI.

### Out of Scope (not implemented)

- Comparing more than two files at once.
- Comparing multiple worksheets within a single run (one sheet per file per run).
- Scheduling or automation (no cron, no CLI batch mode).
- Saving comparison configurations to disk (configuration lives in Streamlit session state only).
- Fuzzy or partial column-name matching (auto-suggestion requires exact name match).
- Cell-level formatting beyond header styling and wrapped text.
- Password-protected or macro-enabled workbooks (`.xlsm`, `.xls`, password-protected `.xlsx`).
- User accounts, authentication, or role-based access.
- Cloud storage or database backends (files are processed in memory / temp files only).

## Current Capabilities

- Upload two `.xlsx` files through the Streamlit UI (50 MB max per file).
- Select one worksheet from each file.
- Configure matching key pairs with auto-suggestion of exact-name matches.
- Configure comparison column pairs (Text / Numeric / Date).
- Run the comparison and view a dashboard with summary metrics.
- Preview results in five tabs: Unchanged Records, Records With Changes, Missing From File B, New Records in File B, Records Needing Attention.
- Download a formatted six-sheet `.xlsx` report.
- Run 303 unit tests with pytest; lint and format with ruff.

## Current Limitations

- Only `.xlsx` files are supported (no `.xls`, `.xlsm`, `.csv`, `.ods`).
- Maximum file size is 50 MB per file.
- One worksheet per file per comparison run.
- No fuzzy column matching; auto-suggestion pairs columns only when names are identical.
- Comparison configuration is not persisted between browser sessions (Streamlit session state only).
- The duplicate-key policy is fixed at `error` (duplicates are reported but do not block matching of unambiguous records); the `all` policy is explicitly unsupported.
- Numeric and date comparison apply only to column pairs explicitly configured as Numeric or Date; all other pairs use text comparison (so `100` vs `100.0` are treated as different text values unless configured as Numeric).
- Date parsing uses `dayfirst=False` only.
- No CLI entry point; the application runs via `streamlit run app/streamlit_app.py`.
- No license has been selected yet (see LICENSE section in README.md).