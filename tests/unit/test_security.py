"""Unit tests for security validation."""

import zipfile
from pathlib import Path

import openpyxl
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
    with pytest.raises(SecurityError, match='not a valid ZIP archive'):
        validate_uploaded_file(not_xlsx)


def test_validate_decompressed_size_within_limit(tmp_path: Path) -> None:
    """A normal valid workbook passes the decompressed-size check."""
    path = tmp_path / 'valid.xlsx'
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Sheet1')
    ws['A1'] = 'hello'
    ws['B1'] = 'world'
    wb.save(path)
    result = validate_uploaded_file(path)
    assert result == path


def test_validate_decompressed_size_exceeded(tmp_path: Path) -> None:
    """A workbook whose declared decompressed size exceeds the limit is rejected."""
    base = tmp_path / 'base.xlsx'
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet('Sheet1')
    ws['A1'] = 'hello'
    wb.save(base)

    # Repackage with a dummy member whose actual uncompressed data exceeds the
    # limit. Highly compressible data (repeated bytes) keeps the compressed
    # size small while the decompressed size is large, simulating a ZIP bomb.
    oversized = tmp_path / 'oversized.xlsx'
    with zipfile.ZipFile(base, 'r') as zin:
        infos = zin.infolist()
        with zipfile.ZipFile(oversized, 'w', zipfile.ZIP_DEFLATED) as zout:
            for info in infos:
                data = zin.read(info.filename)
                zout.writestr(info, data)
            # 250 MB of repeated bytes compresses to well under 50 MB but
            # decompresses to 250 MB (> 200 MB limit).
            dummy = zipfile.ZipInfo('dummy.xml')
            dummy.compress_type = zipfile.ZIP_DEFLATED
            zout.writestr(dummy, b'A' * 250_000_000)

    with pytest.raises(SecurityError, match='decompressed size exceeds maximum'):
        validate_uploaded_file(oversized)
