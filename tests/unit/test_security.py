"""Unit tests for security validation."""

from pathlib import Path

import pytest

from excel_compare.exceptions import SecurityError
from excel_compare.security import validate_uploaded_file


def test_validate_nonexistent_file(tmp_path: Path) -> None:
    with pytest.raises(SecurityError, match='does not exist'):
        validate_uploaded_file(tmp_path / 'missing.xlsx')


def test_validate_directory(tmp_path: Path) -> None:
    with pytest.raises(SecurityError, match='directory'):
        validate_uploaded_file(tmp_path)


def test_validate_unsupported_extension(tmp_path: Path) -> None:
    bad = tmp_path / 'data.csv'
    bad.write_text('a,b\n1,2\n')
    with pytest.raises(SecurityError, match='Unsupported file extension'):
        validate_uploaded_file(bad)


def test_validate_empty_file(tmp_path: Path) -> None:
    empty = tmp_path / 'empty.xlsx'
    empty.write_bytes(b'')
    with pytest.raises(SecurityError, match='empty'):
        validate_uploaded_file(empty)


def test_validate_unreadable_workbook(tmp_path: Path) -> None:
    not_xlsx = tmp_path / 'not_xlsx.xlsx'
    not_xlsx.write_text('this is not an excel file\n')
    with pytest.raises(SecurityError, match='not a readable Excel workbook'):
        validate_uploaded_file(not_xlsx)
