"""Excel file loading.

Reads workbooks into DataFrames, validates required sheets/columns,
and prunes to only the columns needed for comparison.
Original column names, cell values, and data types are preserved.
"""

from __future__ import annotations

import io
from pathlib import Path
from typing import Any, BinaryIO, Dict, List, Optional, Union

import openpyxl
import pandas as pd

from .exceptions import LoadError, SecurityError
from .models import LoadedSheet
from .security import validate_uploaded_file


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
        mapping. Physical row 1 is the header; first data row is row 2.

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

        # Iterate rows to capture header and data with physical row numbers.
        rows_iter = worksheet.iter_rows(values_only=True)

        try:
            header_row = next(rows_iter)
        except StopIteration as exc:
            raise LoadError(f"Worksheet '{sheet_name}' is empty.") from exc  # noqa: TRY003, B904

        headers = [str(h) if h is not None else '' for h in header_row]

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

        # Read data rows, tracking physical Excel row numbers.
        data: List[List[Any]] = []
        excel_row_mapping: Dict[int, int] = {}
        df_index = 0

        # Physical row 1 is the header; data starts at row 2.
        physical_row = 2
        for row_values in rows_iter:
            # Extract selected columns, converting all values to strings
            # to preserve the dtype=str behavior of pandas.read_excel.
            # This keeps leading zeros in text identifiers intact.
            if col_indices:
                row_data = [
                    str(row_values[i])
                    if i < len(row_values) and row_values[i] is not None
                    else None
                    for i in col_indices
                ]
            else:
                row_data = [str(v) if v is not None else None for v in row_values]

            # Check if the row is entirely empty (skip truly blank trailing rows).
            if all(v is None or v == '' for v in row_data):
                physical_row += 1
                continue

            data.append(row_data)
            excel_row_mapping[df_index] = physical_row
            df_index += 1
            physical_row += 1

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
