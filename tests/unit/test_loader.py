"""Unit tests for the Excel loader module."""

from __future__ import annotations

from io import BytesIO

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
    # Leading zeros preserved as text (dtype=str in loader)
    assert df.iloc[0]['id'] == '00123'
    # Original whitespace preserved (loader does not normalize)
    assert df.iloc[0]['name'] == ' Alice '
    # Amount preserved as string (dtype=str in loader)
    assert df.iloc[0]['amount'] == '100'


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
