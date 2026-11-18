"""Shared pytest fixtures."""

from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Any, Dict

import openpyxl
import pytest


@pytest.fixture
def tmp_dir() -> Path:
    td = tempfile.TemporaryDirectory()
    try:
        yield Path(td.name)
    finally:
        try:
            td.cleanup()
        except PermissionError:
            pass


@pytest.fixture
def make_excel_file(tmp_dir: Path):
    """Create an .xlsx workbook from a dict of sheet_name -> list-of-dicts rows.

    Returns the written file path.
    """

    def _create(sheets: Dict[str, list[dict[str, Any]]], filename: str = 'sample.xlsx') -> Path:
        path = tmp_dir / filename
        workbook = openpyxl.Workbook()
        # Remove default sheet
        workbook.remove(workbook.active)
        for sheet_name, rows in sheets.items():
            worksheet = workbook.create_sheet(sheet_name)
            if not rows:
                continue
            headers = list(rows[0].keys())
            worksheet.append(headers)
            for row in rows:
                worksheet.append([row.get(h) for h in headers])
        workbook.save(path)
        return path

    return _create
