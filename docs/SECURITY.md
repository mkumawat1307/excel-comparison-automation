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
| Readability | Must open as a valid workbook via `openpyxl.load_workbook(read_only=True, data_only=True)` | `SecurityError` |

The file is opened read-only with `data_only=True`, so it is never modified during validation.

### 2. Streamlit Upload Validation (`streamlit_app.py`)

`_validate_upload()` performs client-side checks before the security validation:

- Rejects files larger than 50 MB with an `st.error` message.
- Rejects files that do not end in `.xlsx`.
- Calls `validate_uploaded_file()` on a temporary copy for the readability check.

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

### 5. Environment Variables and Secrets

- The application does **not** read environment variables for configuration.
- No API keys, passwords, or secrets are stored in the codebase.
- The `.gitignore` excludes `.env`, `.env.*` (with `!.env.example` allowed), so no secrets file is committed.
- No database connections, cloud storage, or external API calls are made.

### 6. All Processing Is Local

- Comparison runs entirely on the user's machine.
- No data is transmitted over the network.
- No telemetry or reporting to external services.

## Security Tests

`tests/unit/test_security.py` covers:

- `test_validate_nonexistent_file` — `SecurityError` for missing file.
- `test_validate_directory` — `SecurityError` for directory path.
- `test_validate_unsupported_extension` — `SecurityError` for `.csv`.
- `test_validate_empty_file` — `SecurityError` for empty file.
- `test_validate_unreadable_workbook` — `SecurityError` for non-Excel content with `.xlsx` extension.

## Recommended Future Improvements (Not Implemented)

The following are NOT currently implemented. They are listed for awareness only; do not describe them as existing features.

| Area | Recommended improvement | Status |
|------|------------------------|--------|
| Virus scanning | Scan uploaded files with an antivirus engine before processing. | Not implemented |
| File content sanitization | Remove external links, embedded objects, or macros from workbooks before reading. | Not implemented (`.xlsm` already rejected by extension check) |
| Output file integrity | Add checksums or digital signatures to generated reports. | Not implemented |
| Access control | User authentication, role-based access, or audit logging. | Not implemented (single-user local tool) |
| Secure temporary file cleanup | Use `tempfile` with `delete=True` and stricter permissions on temp files. | Partially implemented (temp files are unlinked after use) |
| Content security policy | Streamlit runs without CSP hardening by default. | Not implemented |
| Dependency vulnerability scanning | Automated scanning of `pandas`, `openpyxl`, `streamlit` for known vulnerabilities. | Not implemented |

## Sensitive Data Considerations

- The application processes data entirely in memory and local temp files. No data leaves the machine.
- Generated reports contain the same data as the source files (plus row references). Treat the output report with the same confidentiality as the source Excel files.
- Temp files created during report generation are deleted after the report bytes are read, but no secure overwrite is performed.
- If your source files contain personally identifiable information (PII), financial data, or health information, handle the downloaded report according to your organization's data-handling policies.