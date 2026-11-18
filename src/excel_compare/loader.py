"""Excel file loading.

Reads workbooks into DataFrames, validates required sheets/columns,
and prunes to only the columns needed for comparison.
Original column names, cell values, and data types are preserved.
"""

from __future__ import annotations

import io
from datetime import date, datetime
from pathlib import Path
from typing import Any, BinaryIO, Dict, List, Optional, Union

import openpyxl
import pandas as pd
from openpyxl.utils import get_column_letter

from .config import (
    MAX_WORKSHEET_CELLS,
    MAX_WORKSHEET_COLUMNS,
    MAX_WORKSHEET_ROWS,
)
from .exceptions import LoadError, SecurityError
from .models import LoadedSheet
from .security import validate_uploaded_file


def _zero_padded_width(number_format: Optional[str]) -> Optional[int]:
    """Return zero-padding width for simple integer formats like 00000.

    Returns the integer-part length only when it is non-empty, all zeros,
    and the format has no fractional part (e.g. ``00000`` or ``0000``).
    Formats like ``0.00``, ``00000.00``, or ``#,##0`` return None so that
    native numeric values are preserved.
    """
    if not number_format:
        return None
    first_section = str(number_format).split(';')[0]
    if '.' in first_section:
        return None
    integer_part = first_section.split('.')[0]
    if integer_part and all(ch == '0' for ch in integer_part):
        return len(integer_part)
    return None


def _resolve_cell_value(cell: Any) -> Any:
    """Resolve an openpyxl cell to its preserved Python value.

    Native numbers, dates, datetimes, booleans, and text are preserved.
    A numeric cell with a simple integer zero-padding format (e.g. 00000)
    is returned as its zero-padded text representation so displayed
    identifiers such as 00123 are not lost. Complex or fractional formats
    keep the native numeric value.
    """
    value = cell.value
    if value is None:
        return None

    # Booleans must be checked before ints (bool subclasses int).
    if isinstance(value, bool):
        return value

    # Only integer values are eligible for zero-padding; floats always stay native.
    if isinstance(value, int):
        width = _zero_padded_width(getattr(cell, 'number_format', None))
        if width is not None:
            try:
                if value < 0:
                    # Sign must not count towards the required digit width.
                    return '-' + f'{abs(value):0{width}d}'
                else:
                    return f'{value:0{width}d}'
            except (ValueError, TypeError):
                return str(value)
        return value

    if isinstance(value, float):
        return value

    if isinstance(value, (datetime, date)):
        return value

    if isinstance(value, str):
        return value

    return str(value)


def _validate_headers(headers: List[str], sheet_name: str, data_rows: List[List[Any]]) -> None:
    """Validate worksheet headers.

    Rejects blank headers on columns that contain data, and duplicate
    headers. Completely empty columns (blank header, no data below) are
    ignored. Valid non-blank column names are preserved exactly.

    Raises LoadError with a user-friendly message on the first problem found.
    """
    n_cols = len(headers)

    # Identify data-bearing columns: any cell below the header has data.
    data_bearing = [False] * n_cols
    for row in data_rows:
        for i in range(min(len(row), n_cols)):
            v = row[i] if i < len(row) else None
            if v is not None and v != '':
                data_bearing[i] = True

    # Blank headers on data-bearing columns.
    blank_positions: List[int] = []
    for i, h in enumerate(headers):
        if data_bearing[i] and (h is None or (isinstance(h, str) and h.strip() == '')):
            blank_positions.append(i + 1)  # 1-based column number
    if blank_positions:
        col_refs = []
        for pos in blank_positions:
            letter = get_column_letter(pos)
            col_refs.append(f'column {pos} ({letter})')
        raise LoadError(  # noqa: TRY003
            f"Worksheet '{sheet_name}' has blank column header(s) at "
            f'{", ".join(col_refs)}. All columns containing data must have a header name.'
        )

    # Duplicate headers.
    seen: Dict[str, List[int]] = {}
    for i, h in enumerate(headers):
        seen.setdefault(h, []).append(i + 1)
    dup_messages: List[str] = []
    for h, positions in seen.items():
        if len(positions) > 1:
            col_refs = [get_column_letter(p) for p in positions]
            dup_messages.append(
                f"'{h}' at columns {', '.join(str(p) for p in positions)} ({', '.join(col_refs)})"
            )
    if dup_messages:
        raise LoadError(  # noqa: TRY003
            f"Worksheet '{sheet_name}' has duplicate column header(s): "
            f'{"; ".join(dup_messages)}. Column headers must be unique.'
        )


def list_sheets(file_source: Union[str, Path, BinaryIO]) -> list[str]:
    """Return worksheet names in the workbook.

    Parameters
    ----------
    file_source : str | Path | BinaryIO
        Path to an .xlsx file, or a file-like object (e.g. Streamlit
        UploadedFile or BytesIO). The object's position is reset before
        reading so repeated calls work safely.

    Returns
    -------
    list[str]
        Worksheet names in the workbook.

    Raises
    ------
    LoadError
        If the file cannot be read as a workbook or contains no sheets.
    """
    workbook = _open_workbook(file_source)
    try:
        sheet_names = workbook.sheetnames
        if not sheet_names:
            raise LoadError('Workbook contains no worksheets.')  # noqa: TRY003
        return list(sheet_names)
    finally:
        workbook.close()
        _reset_position(file_source)


def load_sheet(
    file_source: Union[str, Path, BinaryIO],
    sheet_name: str,
    columns: Optional[list[str]] = None,
    required_columns: Optional[list[str]] = None,
) -> pd.DataFrame:
    """Load a single worksheet into a DataFrame.

    Original column names, cell values, and data types are preserved.
    The source object is never modified.

    Parameters
    ----------
    file_source : str | Path | BinaryIO
        Path to an .xlsx file, or a file-like object.
    sheet_name : str
        Name of the worksheet to load.
    columns : list[str] | None
        If provided, only these columns are read (reduces memory usage).
    required_columns : list[str] | None
        If provided, all listed columns must exist in the sheet.

    Returns
    -------
    pd.DataFrame
        The worksheet contents with original types preserved.

    Raises
    ------
    LoadError
        If the file is invalid, the sheet is missing, the sheet/workbook
        is empty, or required columns are absent.
    """
    loaded = load_sheet_with_metadata(
        file_source, sheet_name, columns=columns, required_columns=required_columns
    )
    return loaded.df


def load_sheet_with_metadata(
    file_source: Union[str, Path, BinaryIO],
    sheet_name: str,
    columns: Optional[list[str]] = None,
    required_columns: Optional[list[str]] = None,
    track_excel_row: bool = True,
) -> LoadedSheet:
    """Load a worksheet into a LoadedSheet with Excel row-number metadata.

    The original DataFrame retains exactly the user's columns and values.
    Excel row numbers are tracked in ``excel_row_mapping`` and never
    mutate the DataFrame.

    Parameters
    ----------
    file_source : str | Path | BinaryIO
        Path to an .xlsx file, or a file-like object.
    sheet_name : str
        Name of the worksheet to load.
    columns : list[str] | None
        If provided, only these columns are read.
    required_columns : list[str] | None
        If provided, all listed columns must exist in the sheet.
    track_excel_row : bool
        If True, build the DataFrame index -> physical Excel row number
        mapping. Row numbers are the worksheet's actual physical rows, so
        the header may be any row and leading blank rows or interior gaps
        are preserved.

    Returns
    -------
    LoadedSheet
        Container with the DataFrame and row-number mapping.

    Raises
    ------
    LoadError
        If the file is invalid, the sheet is missing, the sheet/workbook
        is empty, or required columns are absent.
    """
    # Resolve path inputs through security validation first.
    if isinstance(file_source, (str, Path)):
        try:
            path = validate_uploaded_file(file_source)
        except SecurityError as exc:
            raise LoadError(f'Failed to read file: {exc}') from exc  # noqa: TRY003
        file_obj: Union[str, Path, BinaryIO] = path
    else:
        file_obj = file_source

    available = list_sheets(file_obj)
    if sheet_name not in available:
        raise LoadError(  # noqa: TRY003
            f"Worksheet '{sheet_name}' not found. Available worksheets: {', '.join(available)}"
        )

    # Read with openpyxl directly to preserve physical row-number mapping.
    workbook = _open_workbook(file_obj)
    try:
        if sheet_name not in workbook.sheetnames:
            raise LoadError(  # noqa: TRY003
                f"Worksheet '{sheet_name}' not found. "
                f'Available worksheets: {", ".join(workbook.sheetnames)}'
            )

        worksheet = workbook[sheet_name]

        # Reset dimension metadata to neutralize potentially inflated or
        # underreported worksheet dimensions from the XML <dimension> element.
        # After reset, iter_rows() yields all XML rows and sizes each row tuple
        # to its actual last cell, so counters below enforce limits on real data.
        worksheet.reset_dimensions()

        # Iterate rows as cell objects to preserve values and number formats.
        rows_iter = worksheet.iter_rows()

        # One counter bounds every yielded physical row: leading empty rows,
        # the header, interior gaps, and data rows. reset_dimensions() makes
        # iter_rows() yield one tuple per physical worksheet row starting at
        # row 1, so this counter is also the physical Excel row number.
        rows_yielded = 0
        physical_row = 0

        # Read all data rows, resolving cell values and tracking physical row numbers.
        all_row_data: List[List[Any]] = []
        excel_row_mapping: Dict[int, int] = {}
        df_index = 0

        # Resource counters enforced during row iteration.
        actual_rows = 0
        actual_columns = 0
        actual_cells = 0

        # Locate the header row, skipping leading empty rows. Every yielded row
        # counts toward the row limit so leading blank rows cannot bypass it.
        header_row = None
        while True:
            try:
                header_row = next(rows_iter)
            except StopIteration as exc:
                raise LoadError(f"Worksheet '{sheet_name}' is empty.") from exc  # noqa: TRY003, B904
            rows_yielded += 1
            physical_row += 1
            if rows_yielded > MAX_WORKSHEET_ROWS:
                raise LoadError(  # noqa: TRY003
                    f"Worksheet '{sheet_name}' exceeds maximum of {MAX_WORKSHEET_ROWS} rows."
                )
            if header_row and any(
                cell.value is not None and cell.value != '' for cell in header_row
            ):
                break

        # Enforce the column limit on the header row as well as data rows.
        if len(header_row) > MAX_WORKSHEET_COLUMNS:
            raise LoadError(  # noqa: TRY003
                f"Worksheet '{sheet_name}' exceeds maximum of {MAX_WORKSHEET_COLUMNS} columns."
            )

        headers = [str(h.value) if h.value is not None else '' for h in header_row]

        for row_cells in rows_iter:
            # Count every yielded row, including empty rows, so leading blanks,
            # the header, and interior gaps share one row budget.
            rows_yielded += 1
            physical_row += 1
            if rows_yielded > MAX_WORKSHEET_ROWS:
                raise LoadError(  # noqa: TRY003
                    f"Worksheet '{sheet_name}' exceeds maximum of {MAX_WORKSHEET_ROWS} rows."
                )

            row_data = [_resolve_cell_value(cell) for cell in row_cells]

            # Enforce maximum physical row width before processing,
            # including formatting-only cells that carry no data value.
            if len(row_cells) > MAX_WORKSHEET_COLUMNS:
                raise LoadError(  # noqa: TRY003
                    f"Worksheet '{sheet_name}' exceeds maximum of {MAX_WORKSHEET_COLUMNS} columns."
                )

            # Check if the row is entirely empty (skip truly blank rows).
            if all(v is None or v == '' for v in row_data):
                continue

            # Count actual data rows and cells before appending to memory.
            actual_rows += 1
            actual_cells += len(row_cells)

            # Track maximum column index with a non-empty value.
            for i, v in enumerate(row_data):
                if v is not None and v != '':
                    actual_columns = max(actual_columns, i + 1)

            # Enforce limits on actual data before retaining the row.
            if actual_rows > MAX_WORKSHEET_ROWS:
                raise LoadError(  # noqa: TRY003
                    f"Worksheet '{sheet_name}' exceeds maximum of {MAX_WORKSHEET_ROWS} rows."
                )
            if actual_columns > MAX_WORKSHEET_COLUMNS:
                raise LoadError(  # noqa: TRY003
                    f"Worksheet '{sheet_name}' exceeds maximum of {MAX_WORKSHEET_COLUMNS} columns."
                )
            if actual_cells > MAX_WORKSHEET_CELLS:
                raise LoadError(  # noqa: TRY003
                    f"Worksheet '{sheet_name}' exceeds maximum of {MAX_WORKSHEET_CELLS} cells."
                )

            all_row_data.append(row_data)
            excel_row_mapping[df_index] = physical_row
            df_index += 1

        # Validate headers before applying column selection.
        _validate_headers(headers, sheet_name, all_row_data)

        # Apply column selection if requested.
        if columns is not None:
            col_indices = []
            selected_headers = []
            for i, h in enumerate(headers):
                if h in columns:
                    col_indices.append(i)
                    selected_headers.append(h)
            missing = [c for c in columns if c not in headers]
            if missing:
                raise LoadError(  # noqa: TRY003
                    f"Missing requested columns in worksheet '{sheet_name}': "
                    f'{", ".join(missing)}. Available: {", ".join(headers)}'
                )
            headers = selected_headers
        else:
            col_indices = list(range(len(headers)))

        # Validate required columns.
        if required_columns:
            missing = [c for c in required_columns if c not in headers]
            if missing:
                raise LoadError(  # noqa: TRY003
                    f"Missing required columns in worksheet '{sheet_name}': "
                    f'{", ".join(missing)}. Available: {", ".join(map(str, headers))}'
                )

        # Build final data using column selection.
        data: List[List[Any]] = []
        if col_indices:
            for row_data in all_row_data:
                data.append([row_data[i] if i < len(row_data) else None for i in col_indices])
        else:
            data = list(all_row_data)

    finally:
        workbook.close()
        _reset_position(file_obj)

    if not data:
        raise LoadError(f"Worksheet '{sheet_name}' is empty.")  # noqa: TRY003

    df = pd.DataFrame(data, columns=headers)

    # Validate required columns exist after DataFrame construction.
    if required_columns:
        missing = [c for c in required_columns if c not in df.columns]
        if missing:
            raise LoadError(  # noqa: TRY003
                f"Missing required columns in worksheet '{sheet_name}': "
                f'{", ".join(missing)}. Available: {", ".join(map(str, df.columns))}'
            )

    return LoadedSheet(
        df=df,
        sheet_name=sheet_name,
        file_source=str(file_source) if isinstance(file_source, (str, Path)) else 'file-like',
        excel_row_mapping=excel_row_mapping,
    )


def _open_workbook(file_source: Union[str, Path, BinaryIO]):
    """Open a workbook with openpyxl in read-only mode.

    File-like objects are read into memory once so they can be reused.
    """
    if isinstance(file_source, (str, Path)):
        path = Path(file_source)
        if not path.exists():
            raise LoadError(f'File does not exist: {path}')  # noqa: TRY003
        return openpyxl.load_workbook(path, read_only=True, data_only=True)

    # File-like object: buffer into BytesIO so position can be reused.
    raw = file_source.read()
    buffer = io.BytesIO(raw)
    _reset_position(file_source)
    return openpyxl.load_workbook(buffer, read_only=True, data_only=True)


def _reset_position(file_source: Union[str, Path, BinaryIO]) -> None:
    """Reset a file-like object's position to 0, if it supports seeking."""
    if isinstance(file_source, (str, Path)):
        return
    if hasattr(file_source, 'seek'):
        try:
            file_source.seek(0)
        except (OSError, io.UnsupportedOperation):
            pass
