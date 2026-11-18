"""Unit tests for the Excel loader module."""

from __future__ import annotations

import re
import zipfile
from datetime import date, datetime
from io import BytesIO
from numbers import Integral, Real
from pathlib import Path

import openpyxl
import pandas as pd
import pytest

from excel_compare.exceptions import LoadError
from excel_compare.loader import list_sheets, load_sheet


def test_load_valid_file(make_excel_file):
    path = make_excel_file(
        {
            'Sheet1': [
                {'id': '1', 'name': 'Alice', 'amount': 100},
                {'id': '2', 'name': 'Bob', 'amount': 200},
            ]
        }
    )
    df = load_sheet(path, 'Sheet1')
    assert len(df) == 2
    assert list(df.columns) == ['id', 'name', 'amount']


def test_load_file_like_object(make_excel_file):
    path = make_excel_file({'Sheet1': [{'id': '1', 'name': 'Alice'}]})
    raw = path.read_bytes()
    buffer = BytesIO(raw)
    df = load_sheet(buffer, 'Sheet1')
    assert len(df) == 1
    # Original buffer position should be reset for reuse
    buffer.seek(0)
    df2 = load_sheet(buffer, 'Sheet1')
    assert len(df2) == 1


def test_list_sheets(make_excel_file):
    path = make_excel_file(
        {
            'Alpha': [{'a': 1}],
            'Beta': [{'b': 2}],
            'Gamma': [{'c': 3}],
        }
    )
    names = list_sheets(path)
    assert set(names) == {'Alpha', 'Beta', 'Gamma'}


def test_list_sheets_file_like(make_excel_file):
    path = make_excel_file({'Sheet1': [{'a': 1}]})
    buffer = BytesIO(path.read_bytes())
    names = list_sheets(buffer)
    assert names == ['Sheet1']


def test_missing_worksheet(make_excel_file):
    path = make_excel_file({'Sheet1': [{'a': 1}]})
    with pytest.raises(LoadError, match='not found'):
        load_sheet(path, 'NoSuchSheet')


def test_empty_worksheet(make_excel_file):
    path = make_excel_file({'Empty': []})
    with pytest.raises(LoadError, match='is empty'):
        load_sheet(path, 'Empty')


def test_missing_required_columns(make_excel_file):
    path = make_excel_file({'Sheet1': [{'id': '1', 'name': 'Alice'}]})
    with pytest.raises(LoadError, match='Missing required columns'):
        load_sheet(path, 'Sheet1', required_columns=['id', 'amount'])


def test_column_selection(make_excel_file):
    path = make_excel_file({'Sheet1': [{'id': '1', 'name': 'Alice', 'amount': 100}]})
    df = load_sheet(path, 'Sheet1', columns=['id', 'amount'])
    assert list(df.columns) == ['id', 'amount']
    assert 'name' not in df.columns


def test_original_values_preserved(make_excel_file):
    path = make_excel_file(
        {
            'Sheet1': [
                {'id': '00123', 'name': ' Alice ', 'amount': 100},
            ]
        }
    )
    df = load_sheet(path, 'Sheet1')
    # Leading zeros preserved as text (text cells remain str)
    assert df.iloc[0]['id'] == '00123'
    # Original whitespace preserved (loader does not normalize)
    assert df.iloc[0]['name'] == ' Alice '
    # Amount preserved as native integer (not converted to str)
    assert df.iloc[0]['amount'] == 100
    assert isinstance(df.iloc[0]['amount'], Integral)


def test_invalid_file(tmp_dir):
    bad = tmp_dir / 'bad.xlsx'
    bad.write_text('this is not an excel file')
    with pytest.raises(LoadError, match='Failed to read'):
        load_sheet(bad, 'Sheet1')


def test_nonexistent_file(tmp_dir):
    with pytest.raises(LoadError, match='does not exist'):
        load_sheet(tmp_dir / 'missing.xlsx', 'Sheet1')


def test_repeated_reads_file_like(make_excel_file):
    path = make_excel_file({'Sheet1': [{'a': 1}]})
    buffer = BytesIO(path.read_bytes())
    # First read
    df1 = load_sheet(buffer, 'Sheet1')
    # Second read on same buffer should still work (position reset)
    df2 = load_sheet(buffer, 'Sheet1')
    pd.testing.assert_frame_equal(df1, df2)


# ---------------------------------------------------------------------------
# Data type preservation regression tests (Issue 1)
# ---------------------------------------------------------------------------


def _make_typed_workbook(tmp_dir: Path, rows: list, filename: str = 'types.xlsx') -> Path:
    """Create an .xlsx workbook writing rows directly to cells (preserves types)."""
    path = tmp_dir / filename
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Sheet1')
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)
    wb.save(path)
    return path


def test_integer_cell_preserved_as_int(tmp_dir: Path):
    path = _make_typed_workbook(tmp_dir, [['id', 'amount'], [1, 100]])
    df = load_sheet(path, 'Sheet1')
    assert df.iloc[0]['id'] == 1
    assert isinstance(df.iloc[0]['id'], Integral)
    assert df.iloc[0]['amount'] == 100
    assert isinstance(df.iloc[0]['amount'], Integral)


def test_float_cell_preserved_as_float(tmp_dir: Path):
    path = _make_typed_workbook(tmp_dir, [['amount'], [100.5]])
    df = load_sheet(path, 'Sheet1')
    assert df.iloc[0]['amount'] == 100.5
    assert isinstance(df.iloc[0]['amount'], float)


def test_date_cell_preserved_as_date(tmp_dir: Path):
    path = _make_typed_workbook(tmp_dir, [['d'], [date(2024, 1, 15)]])
    df = load_sheet(path, 'Sheet1')
    val = df.iloc[0]['d']
    # Pandas stores dates as Timestamp; value preserves the date.
    assert hasattr(val, 'date')
    assert val.date() == date(2024, 1, 15)


def test_datetime_cell_preserved_as_datetime(tmp_dir: Path):
    path = _make_typed_workbook(tmp_dir, [['dt'], [datetime(2024, 1, 15, 9, 30, 0)]])
    df = load_sheet(path, 'Sheet1')
    val = df.iloc[0]['dt']
    # Pandas stores datetimes as Timestamp; value preserves the datetime.
    assert hasattr(val, 'to_pydatetime')
    assert val.to_pydatetime() == datetime(2024, 1, 15, 9, 30, 0)


def test_boolean_cell_preserved_as_bool(tmp_dir: Path):
    path = _make_typed_workbook(tmp_dir, [['flag'], [True]])
    df = load_sheet(path, 'Sheet1')
    assert df.iloc[0]['flag']
    assert bool(df.iloc[0]['flag']) is True


def test_text_identifier_leading_zeros_preserved(tmp_dir: Path):
    path = _make_typed_workbook(tmp_dir, [['id'], ['00123']])
    df = load_sheet(path, 'Sheet1')
    assert df.iloc[0]['id'] == '00123'
    assert isinstance(df.iloc[0]['id'], str)


def test_zero_padded_format_00000_returns_zero_padded_text(tmp_dir: Path):
    """Numeric 123 with format 00000 should return '00123' as text."""
    path = tmp_dir / 'zeropad.xlsx'
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Sheet1')
    ws['A1'] = 'id'
    ws['A2'] = 123
    ws['A2'].number_format = '00000'
    wb.save(path)

    df = load_sheet(path, 'Sheet1')
    assert df.iloc[0]['id'] == '00123'
    assert isinstance(df.iloc[0]['id'], str)


def test_zero_padded_format_0000_returns_zero_padded_text(tmp_dir: Path):
    """Numeric 123 with format 0000 should return '0123' as text."""
    path = tmp_dir / 'zeropad.xlsx'
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Sheet1')
    ws['A1'] = 'id'
    ws['A2'] = 123
    ws['A2'].number_format = '0000'
    wb.save(path)

    df = load_sheet(path, 'Sheet1')
    assert df.iloc[0]['id'] == '0123'
    assert isinstance(df.iloc[0]['id'], str)


def test_format_0_00_keeps_native_number(tmp_dir: Path):
    """Format 0.00 has a fractional part -> native number preserved, not zero-padded."""
    path = tmp_dir / 'fmt.xlsx'
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Sheet1')
    ws['A1'] = 'amount'
    ws['A2'] = 123
    ws['A2'].number_format = '0.00'
    wb.save(path)

    df = load_sheet(path, 'Sheet1')
    assert df.iloc[0]['amount'] == 123
    assert isinstance(df.iloc[0]['amount'], Integral)


def test_format_00000_00_keeps_native_number(tmp_dir: Path):
    """Format 00000.00 has a fractional part -> native number preserved."""
    path = tmp_dir / 'fmt.xlsx'
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Sheet1')
    ws['A1'] = 'amount'
    ws['A2'] = 123
    ws['A2'].number_format = '00000.00'
    wb.save(path)

    df = load_sheet(path, 'Sheet1')
    assert df.iloc[0]['amount'] == 123
    assert isinstance(df.iloc[0]['amount'], Integral)


def test_negative_value_zero_padded_format(tmp_dir: Path):
    """Negative numeric with zero-padding format returns padded text with sign preserved."""
    path = tmp_dir / 'neg.xlsx'
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Sheet1')
    ws['A1'] = 'id'
    ws['A2'] = -123
    ws['A2'].number_format = '00000'
    wb.save(path)

    df = load_sheet(path, 'Sheet1')
    assert df.iloc[0]['id'] == '-00123'
    assert isinstance(df.iloc[0]['id'], str)


def test_decimal_value_zero_padded_format_keeps_native(tmp_dir: Path):
    """Decimal numeric with zero-padding integer format keeps native float value."""
    path = tmp_dir / 'dec.xlsx'
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Sheet1')
    ws['A1'] = 'amount'
    ws['A2'] = 123.5
    ws['A2'].number_format = '00000'
    wb.save(path)

    df = load_sheet(path, 'Sheet1')
    assert df.iloc[0]['amount'] == 123.5
    assert isinstance(df.iloc[0]['amount'], Real)


# ---------------------------------------------------------------------------
# Header validation regression tests (Issue 2)
# ---------------------------------------------------------------------------


def test_blank_header_on_data_column_rejected(tmp_dir: Path):
    """Blank header on a column that contains data must raise LoadError."""
    path = tmp_dir / 'blank.xlsx'
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Sheet1')
    ws['A1'] = 'id'
    ws['B1'] = ''  # blank header on column B
    ws['A2'] = 1
    ws['B2'] = 'data'  # column B has data
    wb.save(path)

    with pytest.raises(LoadError, match='blank column header'):
        load_sheet(path, 'Sheet1')


def test_duplicate_headers_rejected(tmp_dir: Path):
    """Duplicate column headers must raise LoadError."""
    path = tmp_dir / 'dup.xlsx'
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Sheet1')
    ws['A1'] = 'id'
    ws['B1'] = 'id'  # duplicate
    ws['A2'] = 1
    ws['B2'] = 2
    wb.save(path)

    with pytest.raises(LoadError, match='duplicate column header'):
        load_sheet(path, 'Sheet1')


def test_valid_headers_loaded_exactly(tmp_dir: Path):
    """Valid unique non-blank headers are preserved exactly."""
    path = _make_typed_workbook(tmp_dir, [['ID', 'Name', 'Amount'], [1, 'Alice', 100]])
    df = load_sheet(path, 'Sheet1')
    assert list(df.columns) == ['ID', 'Name', 'Amount']
    assert len(df) == 1


def test_blank_header_no_data_ignored(tmp_dir: Path):
    """Completely empty column (blank header, no data) is ignored, not rejected."""
    path = tmp_dir / 'empty_col.xlsx'
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Sheet1')
    ws['A1'] = 'id'
    ws['B1'] = ''  # blank header, no data below
    ws['A2'] = 1
    wb.save(path)

    df = load_sheet(path, 'Sheet1')
    assert 'id' in df.columns
    assert len(df) == 1


# ---------------------------------------------------------------------------
# Worksheet dimension and cell-count protection regression tests (Issue 6)
# Use monkeypatched limits for efficiency; no enormous physical files.
# ---------------------------------------------------------------------------


def test_load_worksheet_exceeds_max_rows(tmp_dir: Path, monkeypatch):
    """Rows beyond MAX_WORKSHEET_ROWS are rejected."""
    monkeypatch.setattr('excel_compare.loader.MAX_WORKSHEET_ROWS', 5)
    path = tmp_dir / 'too_many_rows.xlsx'
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Sheet1')
    ws['A1'] = 'id'
    for i in range(2, 8):
        ws.cell(row=i, column=1, value=f'row{i}')
    wb.save(path)

    with pytest.raises(LoadError, match='exceeds maximum of 5 rows'):
        load_sheet(path, 'Sheet1')


def test_load_worksheet_exceeds_max_columns(tmp_dir: Path, monkeypatch):
    """Columns beyond MAX_WORKSHEET_COLUMNS are rejected."""
    monkeypatch.setattr('excel_compare.loader.MAX_WORKSHEET_COLUMNS', 2)
    path = tmp_dir / 'too_many_cols.xlsx'
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Sheet1')
    ws['A1'] = 'a'
    ws['B1'] = 'b'
    ws['C1'] = 'c'
    ws['A2'] = 1
    ws['B2'] = 2
    ws['C2'] = 3
    wb.save(path)

    with pytest.raises(LoadError, match='exceeds maximum of 2 columns'):
        load_sheet(path, 'Sheet1')


def test_load_worksheet_exceeds_max_cells(tmp_dir: Path, monkeypatch):
    """Cells beyond MAX_WORKSHEET_CELLS are rejected."""
    monkeypatch.setattr('excel_compare.loader.MAX_WORKSHEET_CELLS', 5)
    path = tmp_dir / 'too_many_cells.xlsx'
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Sheet1')
    ws['A1'] = 'a'
    ws['B1'] = 'b'
    ws['C1'] = 'c'
    ws['D1'] = 'd'
    ws['E1'] = 'e'
    ws['F1'] = 'f'
    ws['A2'] = 1
    ws['B2'] = 2
    ws['C2'] = 3
    ws['D2'] = 4
    ws['E2'] = 5
    ws['F2'] = 6
    wb.save(path)

    with pytest.raises(LoadError, match='exceeds maximum of 5 cells'):
        load_sheet(path, 'Sheet1')


# Matches the <dimension .../> element regardless of attribute spacing or
# self-closing style. openpyxl serializes it as e.g. <dimension ref="A1:B2" />
# (note the space before "/>"), so a literal no-space replacement would miss.
_DIMENSION_ELEMENT_RE = re.compile(r'<dimension\b[^>]*/>')


def _read_sheet_xml(path: Path) -> str:
    """Return the decoded sheet1.xml from an .xlsx."""
    with zipfile.ZipFile(path, 'r') as zin:
        return zin.read('xl/worksheets/sheet1.xml').decode()


def _rewrite_sheet_xml(path: Path, transform) -> None:
    """Rewrite sheet1.xml in an .xlsx using transform(sheet_xml) -> new_xml.

    Raises AssertionError if the transform did not change the XML, so a test
    cannot silently pass because its replacement failed to match the
    serialized form written by openpyxl.
    """
    with zipfile.ZipFile(path, 'r') as zin:
        entries = {item.filename: zin.read(item.filename) for item in zin.infolist()}
    sheet_xml = entries['xl/worksheets/sheet1.xml'].decode()
    new_xml = transform(sheet_xml)
    assert new_xml != sheet_xml, 'sheet1.xml transform did not modify the XML'
    entries['xl/worksheets/sheet1.xml'] = new_xml.encode()
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for filename, data in entries.items():
            zout.writestr(filename, data)


def test_load_worksheet_inflated_dimension_small_data(tmp_dir: Path, monkeypatch):
    """Inflated dimension metadata does not reject a small legitimate worksheet."""
    monkeypatch.setattr('excel_compare.loader.MAX_WORKSHEET_ROWS', 10)
    monkeypatch.setattr('excel_compare.loader.MAX_WORKSHEET_COLUMNS', 10)
    monkeypatch.setattr('excel_compare.loader.MAX_WORKSHEET_CELLS', 100)

    path = tmp_dir / 'inflated_dim.xlsx'
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Sheet1')
    ws['A1'] = 'h1'
    ws['B1'] = 'h2'
    ws['A2'] = 'hello'
    ws['B2'] = 'world'
    wb.save(path)

    _rewrite_sheet_xml(
        path,
        lambda xml: _DIMENSION_ELEMENT_RE.sub('<dimension ref="A1:XFD1048576"/>', xml),
    )
    # Confirm the declared dimension really is inflated before loading.
    assert 'ref="A1:XFD1048576"' in _read_sheet_xml(path)

    df = load_sheet(path, 'Sheet1')
    assert len(df) == 1
    assert df.iloc[0]['h1'] == 'hello'


def test_load_worksheet_underreported_dimension(tmp_dir: Path, monkeypatch):
    """Underreported dimension does not hide actual data beyond it."""
    monkeypatch.setattr('excel_compare.loader.MAX_WORKSHEET_ROWS', 10)
    monkeypatch.setattr('excel_compare.loader.MAX_WORKSHEET_COLUMNS', 10)
    monkeypatch.setattr('excel_compare.loader.MAX_WORKSHEET_CELLS', 100)

    path = tmp_dir / 'underreported_dim.xlsx'
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Sheet1')
    ws['A1'] = 'h1'
    ws['A6'] = 'fifth'
    wb.save(path)

    _rewrite_sheet_xml(
        path,
        lambda xml: _DIMENSION_ELEMENT_RE.sub('<dimension ref="A1:A1"/>', xml),
    )
    # Confirm the declared dimension understates the real extent (row 1 only).
    assert 'ref="A1:A1"' in _read_sheet_xml(path)

    df = load_sheet(path, 'Sheet1')
    assert len(df) == 1
    assert df.iloc[0]['h1'] == 'fifth'


def test_load_worksheet_missing_dimension(tmp_dir: Path, monkeypatch):
    """Missing dimension element does not prevent loading small worksheets."""
    monkeypatch.setattr('excel_compare.loader.MAX_WORKSHEET_ROWS', 10)
    monkeypatch.setattr('excel_compare.loader.MAX_WORKSHEET_COLUMNS', 10)
    monkeypatch.setattr('excel_compare.loader.MAX_WORKSHEET_CELLS', 100)

    path = tmp_dir / 'no_dim.xlsx'
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Sheet1')
    ws['A1'] = 'hello'
    ws['A2'] = 'world'
    wb.save(path)

    _rewrite_sheet_xml(path, lambda xml: _DIMENSION_ELEMENT_RE.sub('', xml))
    # Confirm the dimension element was actually removed.
    assert '<dimension' not in _read_sheet_xml(path)

    df = load_sheet(path, 'Sheet1')
    assert len(df) == 1
    assert df.iloc[0]['hello'] == 'world'


def test_load_worksheet_large_empty_row_gaps(tmp_dir: Path, monkeypatch):
    """Large empty row gaps between data rows are bounded by row limit."""
    monkeypatch.setattr('excel_compare.loader.MAX_WORKSHEET_ROWS', 50000)
    path = tmp_dir / 'empty_row_gaps.xlsx'
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Sheet1')
    ws['A2'] = 'first'
    ws['A100001'] = 'last'
    wb.save(path)

    with pytest.raises(LoadError, match='exceeds maximum of 50000 rows'):
        load_sheet(path, 'Sheet1')


def test_load_worksheet_formatting_only_row(tmp_dir: Path, monkeypatch):
    """Formatting-only cells count toward row width but not actual data rows."""
    monkeypatch.setattr('excel_compare.loader.MAX_WORKSHEET_ROWS', 10)
    monkeypatch.setattr('excel_compare.loader.MAX_WORKSHEET_COLUMNS', 10)
    monkeypatch.setattr('excel_compare.loader.MAX_WORKSHEET_CELLS', 100)

    path = tmp_dir / 'fmt_only.xlsx'
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Sheet1')
    ws['A1'] = 'h1'
    ws['B1'] = 'h2'
    ws['A2'] = 'data'
    # Formatting-only cell: style but no value
    ws['B2'] = None
    ws['B2'].font = openpyxl.styles.Font(bold=True)
    wb.save(path)

    df = load_sheet(path, 'Sheet1')
    assert len(df) == 1
    assert df.iloc[0]['h1'] == 'data'


def test_load_worksheet_within_limits(tmp_dir: Path, monkeypatch):
    """Normal worksheet within all limits loads successfully."""
    monkeypatch.setattr('excel_compare.loader.MAX_WORKSHEET_ROWS', 1000)
    monkeypatch.setattr('excel_compare.loader.MAX_WORKSHEET_COLUMNS', 100)
    monkeypatch.setattr('excel_compare.loader.MAX_WORKSHEET_CELLS', 100000)
    path = _make_typed_workbook(tmp_dir, [['ID', 'Name', 'Amount'], [1, 'Alice', 100]])
    df = load_sheet(path, 'Sheet1')
    assert list(df.columns) == ['ID', 'Name', 'Amount']
    assert len(df) == 1
