"""Excel report writer.

Writes the ComparisonResult to a formatted .xlsx file with six sheets:
Comparison Overview, Unchanged Records, Records With Changes,
Missing From File B, New Records in File B, Records Needing Attention.

All untrusted text values are written as explicit OpenPyXL text cells
(data_type = 'str') so that strings beginning with =, +, - or @ are never
interpreted as Excel formulas. Original string values are preserved
exactly; no apostrophe or tab prefix is added.
"""

from __future__ import annotations

from typing import Any, List, Optional

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook

from .config import REPORT_SHEETS
from .models import (
    ComparisonConfig,
    ComparisonResult,
    RowDiff,
)
from .reporter import (
    file_a_column_header,
    file_b_column_header,
    summary_label,
)

# Styling constants.
HEADER_FILL = PatternFill(start_color='FF4472C4', end_color='FF4472C4', fill_type='solid')
HEADER_FONT = Font(color='FFFFFFFF', bold=True)
MATCH_FILL = PatternFill(start_color='FFC6EFCE', end_color='FFC6EFCE', fill_type='solid')
MISMATCH_FILL = PatternFill(start_color='FFFFC7CE', end_color='FFFFC7CE', fill_type='solid')
AMBER_FILL = PatternFill(start_color='FFFFEB9C', end_color='FFFFEB9C', fill_type='solid')
GREY_FILL = PatternFill(start_color='FFE7E6E6', end_color='FFE7E6E6', fill_type='solid')
WRAP_ALIGNMENT = Alignment(wrap_text=True, vertical='top')


def _write_text_cell(worksheet, row_num: int, col_num: int, value: Any) -> None:
    """Write a value to a cell as explicit text.

    All string cells are assigned data_type = 's' (OpenPyXL TYPE_STRING) so
    OpenPyXL and Excel treat them as literal text. Strings beginning with
    =, +, - or @ are therefore never executable formulas. The original
    string value is preserved exactly; no apostrophe or tab prefix is added.
    Cells use wrapped text so long values remain readable.
    """
    cell = worksheet.cell(row=row_num, column=col_num)
    cell.value = value
    if isinstance(value, str):
        cell.data_type = 's'
    cell.alignment = WRAP_ALIGNMENT


def _style_header_row(worksheet, row_num: int, column_count: int) -> None:
    """Apply header styling to the first row."""
    for col in range(1, column_count + 1):
        cell = worksheet.cell(row=row_num, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical='center')


def _autofit_columns(worksheet, max_width: int = 60) -> None:
    """Best-fit column widths based on content."""
    for col_idx, column in enumerate(worksheet.columns, start=1):
        max_len = 0
        for cell in column:
            try:
                value = str(cell.value) if cell.value is not None else ''
                max_len = max(max_len, len(value))
            except Exception:
                pass
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(
            max(max_len + 2, 8), max_width
        )


def _write_summary_sheet(workbook: Workbook, result: ComparisonResult) -> None:
    """Write the Comparison Overview sheet."""
    ws = workbook.create_sheet(REPORT_SHEETS[0])
    ws.append(['Metric', 'Value'])
    _style_header_row(ws, 1, 2)

    row_num = 2
    for key, value in result.summary.items():
        display_key = summary_label(key)
        _write_text_cell(ws, row_num, 1, display_key)
        _write_text_cell(ws, row_num, 2, value)
        row_num += 1

    _autofit_columns(ws)


def _write_matching_records_sheet(
    workbook: Workbook,
    result: ComparisonResult,
    config: ComparisonConfig,
    df_a: Optional[pd.DataFrame] = None,
    df_b: Optional[pd.DataFrame] = None,
) -> None:
    """Write the Unchanged Records sheet.

    Only configured columns are shown, in configured order:
    1. Interleaved key column pairs: File A key 1, File B key 1, ...
    2. Selected comparison columns from File B only (values are identical to
       File A for unchanged records, so File A comparison columns are omitted).
    3. Record Status, Excel Row (File A), Excel Row (File B).

    Unmapped columns (neither a key nor a selected comparison column) are not
    displayed.
    """
    ws = workbook.create_sheet(REPORT_SHEETS[1])

    compare_cols_b = [pair.file_b_column for pair in config.compare_pairs]

    # Interleaved key column headers.
    key_headers: list[str] = []
    for pair in config.key_pairs:
        key_headers.append(file_a_column_header(pair.file_a_column))
        key_headers.append(file_b_column_header(pair.file_b_column))

    # File B comparison column headers (File B names only).
    compare_headers_b = [file_b_column_header(c) for c in compare_cols_b]

    columns = (
        key_headers
        + compare_headers_b
        + ['Record Status', 'Excel Row (File A)', 'Excel Row (File B)']
    )
    ws.append(columns)
    _style_header_row(ws, 1, len(columns))

    row_num = 2
    for row in result.matched:
        values: list[Any] = []

        # Interleaved key column values: File A key 1, File B key 1, ...
        for i, _pair in enumerate(config.key_pairs):
            values.append(row.key_values_a[i] if i < len(row.key_values_a) else '')
            values.append(row.key_values_b[i] if i < len(row.key_values_b) else '')

        # File B comparison values (from full source row, File B column names).
        source_row_b = None
        if df_b is not None and row.index_b is not None and int(row.index_b) in df_b.index:
            source_row_b = df_b.loc[int(row.index_b)]
        for col in compare_cols_b:
            if source_row_b is not None:
                values.append(source_row_b[col])
            else:
                values.append('')

        values.extend(
            [
                row.status.label,
                row.excel_row_a if row.excel_row_a is not None else '',
                row.excel_row_b if row.excel_row_b is not None else '',
            ]
        )
        for col_idx, val in enumerate(values, start=1):
            _write_text_cell(ws, row_num, col_idx, val)
        row_num += 1

    _autofit_columns(ws)


def _write_mismatched_records_sheet(
    workbook: Workbook, result: ComparisonResult, config: ComparisonConfig
) -> None:
    """Write the Records With Changes sheet."""
    key_headers: list[str] = []
    for pair in config.key_pairs:
        key_headers.append(file_a_column_header(pair.file_a_column))
        key_headers.append(file_b_column_header(pair.file_b_column))
    columns = key_headers + [
        'Column Name (File A)',
        'Previous Value',
        'Column Name (File B)',
        'Updated Value',
        'Excel Row (File A)',
        'Excel Row (File B)',
    ]
    ws = workbook.create_sheet(REPORT_SHEETS[2])
    ws.append(columns)
    _style_header_row(ws, 1, len(columns))

    row_num = 2
    for row in result.mismatched:
        key_values = []
        for i, _pair in enumerate(config.key_pairs):
            key_values.append(row.key_values_a[i] if i < len(row.key_values_a) else '')
            key_values.append(row.key_values_b[i] if i < len(row.key_values_b) else '')
        for cell_diff in row.cell_diffs:
            if not cell_diff.are_equal:
                vals = key_values + [
                    cell_diff.column_a,
                    cell_diff.value_a,
                    cell_diff.column_b,
                    cell_diff.value_b,
                    cell_diff.excel_row_a if cell_diff.excel_row_a is not None else '',
                    cell_diff.excel_row_b if cell_diff.excel_row_b is not None else '',
                ]
                for col_idx, val in enumerate(vals, start=1):
                    _write_text_cell(ws, row_num, col_idx, val)
                row_num += 1

    _autofit_columns(ws)


def _write_only_in_sheet(
    workbook: Workbook,
    sheet_name: str,
    rows: List[RowDiff],
    side: str,
    config: ComparisonConfig,
    df: Optional[pd.DataFrame] = None,
) -> None:
    """Write Missing From File B or New Records in File B sheet.

    Displays the complete original record from the present side only,
    using original column names (no File A / File B prefix), plus the
    Excel row reference. No opposite-side columns are added.
    """

    # Column headers: original source column names for the present side,
    # followed by the Excel row reference.
    if side == 'A':
        source_columns = list(df.columns) if df is not None else []
        excel_row_label = 'Excel Row (File A)'
    else:
        source_columns = list(df.columns) if df is not None else []
        excel_row_label = 'Excel Row (File B)'

    columns = source_columns + [excel_row_label]
    ws = workbook.create_sheet(sheet_name)
    ws.append(columns)
    _style_header_row(ws, 1, len(columns))

    row_num = 2
    for row in rows:
        values: list[Any] = []

        # Complete record from the present side using original column names.
        source_row = None
        if df is not None:
            if side == 'A' and row.index_a is not None and int(row.index_a) in df.index:
                source_row = df.loc[int(row.index_a)]
            elif side == 'B' and row.index_b is not None and int(row.index_b) in df.index:
                source_row = df.loc[int(row.index_b)]

        for col in source_columns:
            if source_row is not None:
                values.append(source_row[col])
            else:
                values.append('')

        idx_val = row.excel_row_a if side == 'A' else row.excel_row_b
        values.append(idx_val if idx_val is not None else '')
        for col_idx, val in enumerate(values, start=1):
            _write_text_cell(ws, row_num, col_idx, val)
        row_num += 1

    _autofit_columns(ws)


def _write_data_quality_sheet(
    workbook: Workbook, result: ComparisonResult, config: ComparisonConfig
) -> None:
    """Write the Records Needing Attention sheet."""
    key_headers: list[str] = []
    for pair in config.key_pairs:
        key_headers.append(file_a_column_header(pair.file_a_column))
        key_headers.append(file_b_column_header(pair.file_b_column))
    columns = ['Problem Found', 'File', 'Excel Row'] + key_headers + ['Related Column Values']
    ws = workbook.create_sheet(REPORT_SHEETS[5])
    ws.append(columns)
    _style_header_row(ws, 1, len(columns))

    n_keys = len(config.key_pairs)
    row_num = 2
    for issue in result.data_quality_issues:
        col_values_str = ' | '.join(
            f'{k}={v!r}' for k, v in sorted(issue.all_column_values.items())
        )
        vals = [
            issue.issue_type,
            issue.source_file,
            issue.excel_row_number if issue.excel_row_number is not None else '',
        ]
        for i in range(n_keys):
            vals.append(issue.key_values_a[i] if i < len(issue.key_values_a) else '')
        for i in range(n_keys):
            vals.append(issue.key_values_b[i] if i < len(issue.key_values_b) else '')
        vals.append(col_values_str)
        for col_idx, val in enumerate(vals, start=1):
            _write_text_cell(ws, row_num, col_idx, val)
        row_num += 1

    _autofit_columns(ws)


def write_report(
    result: ComparisonResult,
    output_path: str,
    config: ComparisonConfig,
    df_a: Optional[pd.DataFrame] = None,
    df_b: Optional[pd.DataFrame] = None,
) -> None:
    """Write the comparison report to a formatted XLSX file.

    The six sheets are: Comparison Overview, Unchanged Records, Records With
    Changes, Missing From File B, New Records in File B, Records Needing
    Attention.

    Parameters
    ----------
    result : ComparisonResult
        The assembled comparison result.
    output_path : str
        Path to write the .xlsx file.
    config : ComparisonConfig
        Comparison settings (used for key column headers).
    df_a : pandas.DataFrame, optional
        Original File A DataFrame, used to write complete records on
        record-level sheets.
    df_b : pandas.DataFrame, optional
        Original File B DataFrame, used to write complete records on
        record-level sheets.
    """
    workbook: Workbook = Workbook()
    try:
        # Remove default sheet.
        workbook.remove(workbook.active)

        _write_summary_sheet(workbook, result)
        _write_matching_records_sheet(workbook, result, config, df_a=df_a, df_b=df_b)
        _write_mismatched_records_sheet(workbook, result, config)
        _write_only_in_sheet(workbook, REPORT_SHEETS[3], result.only_in_a, 'A', config, df=df_a)
        _write_only_in_sheet(workbook, REPORT_SHEETS[4], result.only_in_b, 'B', config, df=df_b)
        _write_data_quality_sheet(workbook, result, config)

        workbook.save(output_path)
    finally:
        workbook.close()
