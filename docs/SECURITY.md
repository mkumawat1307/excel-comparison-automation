# Security

This document describes the security measures implemented in Excel Comparison Automation, and distinguishes them from recommended future improvements.

## Implemented Protections

### 1. File Upload Validation (`security.py`)

Every uploaded file is validated before processing by `validate_uploaded_file()`:

| Check | Behavior | Failure raises |
|-------|----------|----------------|
| File existence | Path must exist | `SecurityError` |
| Not a directory | Path must be a file, not a directory | `SecurityError` |
| Extension | Must be `.xlsx` (case-insensitive) | `SecurityError` |
| Non-empty | File size must be greater than 0 bytes | `SecurityError` |
| Size limit | Maximum 50 MB (`50 * 1024 * 1024` bytes) | `SecurityError` |
| Decompressed size limit | Sum of ZIP central-directory uncompressed sizes must not exceed 200 MB (`200 * 1024 * 1024` bytes) | `SecurityError` |
| Valid ZIP archive | File must parse as a ZIP archive; malformed archives are rejected | `SecurityError` |
| Readability | Must open as a valid workbook via `openpyxl.load_workbook(read_only=True, data_only=True)` | `SecurityError` |

The file is opened read-only with `data_only=True`, so it is never modified during validation.

The decompressed-size guard sums `ZipInfo.file_size` from the archive's central directory before the workbook is opened, bounding ZIP-bomb style expansion. This value is declared metadata from the archive index; it is a fast pre-open guard and does not by itself verify archive integrity or the actual decompressed byte count. It sits behind the 50 MB compressed-size limit. A malformed (non-ZIP) file raises `SecurityError` (`File is not a valid ZIP archive`) at this step, before the readability check.

### 2. Streamlit Upload Validation (`streamlit_app.py`)

`_validate_upload()` performs client-side checks before the security validation:

- Rejects files larger than 50 MB with an `st.error` message.
- Rejects files that do not end in `.xlsx`.
- Calls `validate_uploaded_file()` on a temporary copy for the readability check.
- Deletes the temporary copy in a `finally` block. `validate_uploaded_file()` closes the workbook handle it opens, so the file can be deleted on Windows as well; a failed deletion is logged rather than silently ignored.

### 3. Formula-Injection Protection (`writer.py`)

All untrusted text values written to the Excel report are written as explicit OpenPyXL text cells with `data_type = 's'` (`_write_text_cell()`):

- Strings beginning with `=`, `+`, `-`, or `@` are stored as literal text, never interpreted as Excel formulas.
- The original string value is preserved exactly; no apostrophe or tab prefix is added.
- Protection applies to every sheet: Comparison Overview, Unchanged Records, Records With Changes, Missing From File B, New Records in File B, Records Needing Attention.
- Covered by tests: `test_formula_injection_equal_sign`, `test_formula_injection_plus_sign`, `test_formula_injection_minus_sign`, `test_formula_injection_at_sign`, `test_write_formula_injection_protected_on_mismatched_sheet`, `test_writer_formula_injection_on_key_columns`, `test_writer_formula_injection_still_protected`.

### 4. Original Workbook Protection

- Source Excel files are opened read-only (`read_only=True`).
- The loader never modifies the source DataFrame (`normalize_dataframe()` returns a copy).
- `safe_copy_to_temp()` copies validated files to a temp location; the original is never modified.
- Uploaded file bytes are stored in Streamlit session state and written to a temporary file only for the report-writing step, then cleaned up.
- The workbook opened by `validate_uploaded_file()` for the readability check is closed in a `finally` block; the workbook created by `write_report()` is closed in a `finally` block.

### 5. Environment Variables and Secrets

- The application does **not** read environment variables for configuration.
- No API keys, passwords, or secrets are stored in the codebase.
- The `.gitignore` excludes `.env`, `.env.*` (with `!.env.example` allowed), so no secrets file is committed.
- No database connections, cloud storage, or external API calls are made.

### 6. All Processing Is Local

- Comparison runs entirely on the user's machine.
- No data is transmitted over the network.
- No telemetry or reporting to external services.

### 7. Worksheet Resource Limits (`loader.py`)

To prevent a single worksheet from exhausting memory or CPU, `load_sheet_with_metadata()` enforces hard limits while iterating rows. Limits are defined in `src/excel_compare/config.py`:

| Limit | Constant | Value |
|-------|----------|-------|
| Yielded rows | `MAX_WORKSHEET_ROWS` | 250,000 |
| Actual data rows | `MAX_WORKSHEET_ROWS` | 250,000 |
| Columns | `MAX_WORKSHEET_COLUMNS` | 1,000 |
| Cells | `MAX_WORKSHEET_CELLS` | 5,000,000 |

How content is counted, matching the implementation:

- **Yielded rows** counts every yielded physical worksheet row — leading blank rows, the header, interior gaps, and empty rows — and is checked before a row is processed or retained, including while searching for the header.
- **Actual data rows** counts rows with at least one non-empty value.
- **Columns** bounds the width of the header row and of each yielded data row (including formatting-only cells) and the highest column index holding a non-empty value.
- **Cells** sums the width of each row that contains at least one non-empty value.
- An empty row does not count toward actual data rows or cells. An empty cell inside a data row is preserved as `None` and still contributes to width and cell counts. A formatting-only cell contributes to width and cell counts but does not by itself make a row count as data.

Exceeding any limit raises `LoadError` naming the specific limit. Oversized worksheets are **rejected, never silently truncated**.

### 8. Worksheet Dimension Protection (`loader.py`)

Excel's `<dimension>` element declares a worksheet's used range, but this metadata is not trustworthy: it can be inflated, underreported, or absent. Relying on it would either reject small valid worksheets (inflated) or silently drop real data (underreported/missing).

Before iterating, the loader calls `worksheet.reset_dimensions()`, which clears the cached dimension so `iter_rows()` walks the actual worksheet XML and sizes each row to its true last cell. This neutralizes inflated, underreported, and missing dimensions, after which the resource limits in section 7 are enforced on the real data encountered. A worksheet that exceeds a limit is rejected with a clear `LoadError` rather than being silently truncated.

### 9. Temporary Data Cleanup and Session Reset (`streamlit_app.py`)

Uploaded data and generated report data are held only in the Streamlit session and local temporary files, and are released by explicit user actions:

- **Temp files:** `_validate_upload()` and `_run_comparison()` each write to a `NamedTemporaryFile(delete=False)` and delete it in a `finally` block. Failures to delete are logged (via the module logger), not silently swallowed.
- **Finish & Clear Data (Step 6):** after downloading the report, the user can confirm and clear all data. The report and download button remain available until the user confirms.
- **Start Over / Clear Data (sidebar):** available from every step, also confirmation-gated, to discard an unfinished comparison.
- Both actions call `_clear_session_data()`, which calls `st.session_state.clear()` and then increments an uploader nonce (`_session_nonce`). Incrementing the nonce rotates the file-uploader widget keys so previously uploaded files cannot reappear on the next rerun.

Limitations:

- Clearing drops Python references to the data. It is **not** secure physical memory erasure; copies may remain in the Python heap, in Streamlit internals, or in the operating system until they are reclaimed.
- Streamlit's media-file layer may retain a copy of the rendered download bytes for the session.
- The server cannot detect that a browser download completed, so cleanup is never triggered automatically by the download button.
- While a browser tab stays connected, session state persists indefinitely; after disconnect, Streamlit retains the session for `server.disconnectedSessionTTL` (default 120 seconds) before evicting it. This is a deployment setting, not an application-level cleanup.

## Security Tests

`tests/unit/test_security.py` covers:

- `test_validate_nonexistent_file` — `SecurityError` for missing file.
- `test_validate_directory` — `SecurityError` for directory path.
- `test_validate_unsupported_extension` — `SecurityError` for `.csv`.
- `test_validate_empty_file` — `SecurityError` for empty file.
- `test_validate_unreadable_workbook` — `SecurityError` (`not a valid ZIP archive`) for non-Excel content with an `.xlsx` extension.
- `test_validate_decompressed_size_within_limit` — a normal workbook passes the decompressed-size check.
- `test_validate_decompressed_size_exceeded` — a ZIP-bomb style archive whose declared decompressed size exceeds 200 MB is rejected.
- `test_validate_closes_workbook` — the workbook opened for the readability check is closed.
- `test_validate_releases_file_handle_for_deletion` — after validation the file can be deleted (no leaked file handle).

`tests/unit/test_loader.py` additionally covers the worksheet resource limits and dimension protection:

- `test_load_worksheet_exceeds_max_rows`, `test_load_worksheet_exceeds_max_columns`, `test_load_worksheet_exceeds_max_cells` — each limit raises `LoadError` with the corresponding message.
- `test_load_worksheet_large_empty_row_gaps` — large empty row gaps are bounded by the yielded-row limit.
- `test_load_worksheet_header_exceeds_max_columns`, `test_load_worksheet_header_at_max_columns_allowed` — the column limit also applies to the header row, and a header exactly at the limit loads.
- `test_load_worksheet_leading_empty_rows_exceed_max_rows`, `test_load_worksheet_leading_empty_rows_within_limit` — leading blank rows count toward the yielded-row limit while the header is located.
- `test_excel_row_mapping_header_on_first_row`, `test_excel_row_mapping_with_leading_and_interior_empty_rows` — physical Excel row numbers reflect leading blank rows and interior gaps.
- `test_load_worksheet_formatting_only_row`, `test_load_worksheet_within_limits` — formatting-only cells and within-limit worksheets load as expected.
- `test_load_worksheet_inflated_dimension_small_data`, `test_load_worksheet_underreported_dimension`, `test_load_worksheet_missing_dimension` — inflated, underreported, and missing `<dimension>` metadata do not reject valid data or hide real rows.

## Recommended Future Improvements (Not Implemented)

The following are NOT currently implemented. They are listed for awareness only; do not describe them as existing features.

| Area | Recommended improvement | Status |
|------|------------------------|--------|
| Virus scanning | Scan uploaded files with an antivirus engine before processing. | Not implemented |
| File content sanitization | Remove external links, embedded objects, or macros from workbooks before reading. | Not implemented (`.xlsm` already rejected by extension check) |
| Output file integrity | Add checksums or digital signatures to generated reports. | Not implemented |
| Access control | User authentication, role-based access, or audit logging. | Not implemented (single-user local tool) |
| Secure temporary file cleanup | Use `tempfile` with `delete=True` and stricter permissions on temp files. | Partially implemented (temp files are unlinked in `finally`; failures are logged. No secure overwrite.) |
| Content security policy | Streamlit runs without CSP hardening by default. | Not implemented |
| Dependency vulnerability scanning | Automated scanning of `pandas`, `openpyxl`, `streamlit` for known vulnerabilities. | Not implemented |

## Sensitive Data Considerations

- The application processes data entirely in memory and local temp files. No data leaves the machine.
- Generated reports contain the same data as the source files (plus row references). Treat the output report with the same confidentiality as the source Excel files.
- Temp files created during report generation are deleted after the report bytes are read, but no secure overwrite is performed.
- "Finish & Clear Data" and "Start Over / Clear Data" drop the app's references to uploaded data, loaded worksheets, mappings, comparison results, and report bytes. This is reference release, not secure memory erasure.
- If your source files contain personally identifiable information (PII), financial data, or health information, handle the downloaded report according to your organization's data-handling policies.