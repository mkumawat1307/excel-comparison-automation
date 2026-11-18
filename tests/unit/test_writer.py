"""Unit tests for the writer module (Phase 4)."""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import openpyxl
import pandas as pd
import pytest

import excel_compare.writer as writer_module
from excel_compare.config import REPORT_SHEETS
from excel_compare.models import (
    CellDiff,
    ColumnPair,
    ComparisonConfig,
    ComparisonResult,
    DataQualityIssue,
    DiffStatus,
    DuplicateSourceRecord,
    MatchResult,
    MissingKeyRecord,
    RowDiff,
    SchemaDiff,
)
from excel_compare.writer import write_report


def _config(
    keys=None,
    compare_columns=None,
    numeric_columns=None,
    date_columns=None,
    case_sensitive=False,
    normalize_whitespace=True,
    numeric_tolerance=1e-9,
    key_pairs=None,
    compare_pairs=None,
    numeric_pairs=None,
    date_pairs=None,
) -> ComparisonConfig:
    return ComparisonConfig(
        keys=keys or [],
        compare_columns=compare_columns,
        numeric_columns=numeric_columns,
        date_columns=date_columns,
        key_pairs=key_pairs or [],
        compare_pairs=compare_pairs or [],
        numeric_pairs=numeric_pairs or [],
        date_pairs=date_pairs or [],
        case_sensitive=case_sensitive,
        normalize_whitespace=normalize_whitespace,
        numeric_tolerance=numeric_tolerance,
    )


def _row(data: dict) -> pd.Series:
    return pd.Series(data)


def test_write_report_creates_six_sheets(tmp_path: Path):
    row_a = _row({'id': '1', 'amount': '100'})
    row_b = _row({'id': '1', 'amount': '100'})
    config = _config(keys=['id'], compare_columns=['amount'])

    match_result = MatchResult(
        matched_pairs=[(row_a, row_b)],
        only_in_a_unique=[],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['id', 'amount'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=['amount'],
        ),
    )

    result = ComparisonResult(
        matched=[RowDiff(row_key='id=1', status=DiffStatus.MATCHED)],
        mismatched=[],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'matched_records': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 1},
    )

    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)

    assert output_path.exists()

    workbook = openpyxl.load_workbook(output_path)
    sheet_names = workbook.sheetnames
    assert sheet_names == list(REPORT_SHEETS)
    assert len(sheet_names) == 6
    assert 'Records Needing Attention' in sheet_names
    assert 'Duplicate Records' not in sheet_names


def test_write_report_data_quality_sheet_has_issue_type(tmp_path: Path):
    dup_rec = DuplicateSourceRecord(
        source_file='A',
        original_index=5,
        excel_row_number=7,
        key_values=('1',),
        key_display="id='1'",
        group_occurrence_count=2,
        all_column_values={'id': '1', 'name': 'Alice'},
    )
    missing_rec = MissingKeyRecord(
        source_file='B',
        original_index=3,
        excel_row_number=5,
        key_values=('',),
        key_display="id=''",
        all_column_values={'id': '', 'name': 'Bob'},
    )

    match_result = MatchResult(
        matched_pairs=[],
        only_in_a_unique=[],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[dup_rec],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[missing_rec],
        schema_diff=SchemaDiff(
            common_columns=['id', 'name'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=['name'],
        ),
    )

    result = ComparisonResult(
        matched=[],
        mismatched=[],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[
            DataQualityIssue(
                issue_type='DUPLICATE_KEY',
                source_file=dup_rec.source_file,
                original_index=dup_rec.original_index,
                excel_row_number=dup_rec.excel_row_number,
                key_display=dup_rec.key_display,
                all_column_values=dup_rec.all_column_values,
                key_values_a=dup_rec.key_values,
                key_values_b=(),
            ),
            DataQualityIssue(
                issue_type='MISSING_KEY',
                source_file=missing_rec.source_file,
                original_index=missing_rec.original_index,
                excel_row_number=missing_rec.excel_row_number,
                key_display=missing_rec.key_display,
                all_column_values=missing_rec.all_column_values,
                key_values_a=(),
                key_values_b=missing_rec.key_values,
            ),
        ],
        schema_diff=match_result.schema_diff,
        summary={},
    )

    config = _config(key_pairs=[ColumnPair('id', 'id')], compare_columns=['name'])
    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)

    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Records Needing Attention']

    headers = [cell.value for cell in ws[1]]
    assert 'Problem Found' in headers
    assert 'File' in headers
    assert 'Excel Row' in headers
    assert 'Related Column Values' in headers
    # Per-file key columns should be present (no combined Record ID)
    assert 'Record ID' not in headers
    assert 'File A – id' in headers
    assert 'File B – id' in headers

    issue_types = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        issue_types.append(row[0])
    assert 'DUPLICATE_KEY' in issue_types
    assert 'MISSING_KEY' in issue_types


def test_write_report_preserves_original_values(tmp_path: Path):
    row_a = _row({'id': '00123', 'name': ' Alice ', 'amount': '100'})
    row_b = _row({'id': '00123', 'name': 'alice', 'amount': '100'})
    config = _config(key_pairs=[ColumnPair('id', 'id')], compare_columns=['name', 'amount'])

    match_result = MatchResult(
        matched_pairs=[(row_a, row_b)],
        only_in_a_unique=[],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['id', 'name', 'amount'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=['name', 'amount'],
        ),
    )

    cell_diffs = [
        CellDiff(
            column='name',
            row_key="id='00123'",
            sheet='',
            value_a=' Alice ',
            value_b='alice',
            are_equal=False,
        ),
        CellDiff(
            column='amount',
            row_key="id='00123'",
            sheet='',
            value_a='100',
            value_b='100',
            are_equal=True,
        ),
    ]

    result = ComparisonResult(
        matched=[],
        mismatched=[
            RowDiff(
                row_key="id='00123'",
                status=DiffStatus.MISMATCHED,
                cell_diffs=cell_diffs,
                index_a=0,
                index_b=0,
                key_values_a=('00123',),
                key_values_b=('00123',),
            )
        ],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'mismatched_records': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 1},
    )

    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)

    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Records With Changes']

    found_name_mismatch = False
    for row in ws.iter_rows(min_row=2, values_only=True):
        # Columns: File A.id, File B.id, Field in File A, Value in File A,
        # Field in File B, Value in File B, Excel Row A, Excel Row B
        if row[2] == 'name':
            assert row[3] == ' Alice ', f"Expected ' Alice ', got {row[3]!r}"
            assert row[5] == 'alice', f"Expected 'alice', got {row[5]!r}"
            found_name_mismatch = True
    assert found_name_mismatch, 'Name mismatch row not found in report'


def test_formula_injection_equal_sign(tmp_path: Path):
    """String beginning with '=' is stored as literal text, not formula."""
    row_a = _row({'id': '1', 'name': '=1+1'})
    row_b = _row({'id': '1', 'name': '=1+1'})
    config = _config(keys=['id'], compare_pairs=[ColumnPair('name', 'name')])
    match_result = MatchResult(
        matched_pairs=[(row_a, row_b)],
        only_in_a_unique=[],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['id', 'name'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=['name'],
            exact_name_key_pairs=[ColumnPair('id', 'id')],
            exact_name_compare_pairs=[ColumnPair('name', 'name')],
            renamed_key_pairs=[],
            renamed_compare_pairs=[],
            unmapped_file_a_columns=[],
            unmapped_file_b_columns=[],
        ),
    )
    cell_diffs = [
        [
            CellDiff(
                column_a='name',
                column_b='name',
                row_key="id='1'",
                sheet='',
                value_a='=1+1',
                value_b='=1+1',
                are_equal=True,
            )
        ]
    ]
    result = ComparisonResult(
        matched=[
            RowDiff(
                row_key="id='1'",
                status=DiffStatus.MATCHED,
                cell_diffs=cell_diffs[0],
                index_a=0,
                index_b=0,
            )
        ],
        mismatched=[],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'matched_records': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 1},
    )
    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)
    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Unchanged Records']
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith('='):
                assert cell.value == '=1+1', f"Expected '=1+1', got {cell.value!r}"
                assert cell.data_type == 's', f"Expected 's', got {cell.data_type!r}"


def test_formula_injection_plus_sign(tmp_path: Path):
    """String beginning with '+' is stored as literal text."""
    row_a = _row({'id': '1', 'name': '+100'})
    row_b = _row({'id': '1', 'name': '+100'})
    config = _config(keys=['id'], compare_pairs=[ColumnPair('name', 'name')])
    match_result = MatchResult(
        matched_pairs=[(row_a, row_b)],
        only_in_a_unique=[],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['id', 'name'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=['name'],
            exact_name_key_pairs=[ColumnPair('id', 'id')],
            exact_name_compare_pairs=[ColumnPair('name', 'name')],
            renamed_key_pairs=[],
            renamed_compare_pairs=[],
            unmapped_file_a_columns=[],
            unmapped_file_b_columns=[],
        ),
    )
    cell_diffs = [
        [
            CellDiff(
                column_a='name',
                column_b='name',
                row_key="id='1'",
                sheet='',
                value_a='+100',
                value_b='+100',
                are_equal=True,
            )
        ]
    ]
    result = ComparisonResult(
        matched=[
            RowDiff(
                row_key="id='1'",
                status=DiffStatus.MATCHED,
                cell_diffs=cell_diffs[0],
                index_a=0,
                index_b=0,
            )
        ],
        mismatched=[],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'matched_records': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 1},
    )
    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)
    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Unchanged Records']
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith('+'):
                assert cell.value == '+100', f"Expected '+100', got {cell.value!r}"
                assert cell.data_type == 's', f"Expected 's', got {cell.data_type!r}"


def test_formula_injection_minus_sign(tmp_path: Path):
    """String beginning with '-' is stored as literal text."""
    row_a = _row({'id': '1', 'name': '-50'})
    row_b = _row({'id': '1', 'name': '-50'})
    config = _config(keys=['id'], compare_pairs=[ColumnPair('name', 'name')])
    match_result = MatchResult(
        matched_pairs=[(row_a, row_b)],
        only_in_a_unique=[],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['id', 'name'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=['name'],
            exact_name_key_pairs=[ColumnPair('id', 'id')],
            exact_name_compare_pairs=[ColumnPair('name', 'name')],
            renamed_key_pairs=[],
            renamed_compare_pairs=[],
            unmapped_file_a_columns=[],
            unmapped_file_b_columns=[],
        ),
    )
    cell_diffs = [
        [
            CellDiff(
                column_a='name',
                column_b='name',
                row_key="id='1'",
                sheet='',
                value_a='-50',
                value_b='-50',
                are_equal=True,
            )
        ]
    ]
    result = ComparisonResult(
        matched=[
            RowDiff(
                row_key="id='1'",
                status=DiffStatus.MATCHED,
                cell_diffs=cell_diffs[0],
                index_a=0,
                index_b=0,
            )
        ],
        mismatched=[],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'matched_records': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 1},
    )
    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)
    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Unchanged Records']
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith('-'):
                assert cell.value == '-50', f"Expected '-50', got {cell.value!r}"
                assert cell.data_type == 's', f"Expected 's', got {cell.data_type!r}"


def test_formula_injection_at_sign(tmp_path: Path):
    """String beginning with '@' is stored as literal text."""
    row_a = _row({'id': '1', 'name': '@SUM(A1)'})
    row_b = _row({'id': '1', 'name': '@SUM(A1)'})
    config = _config(keys=['id'], compare_pairs=[ColumnPair('name', 'name')])
    match_result = MatchResult(
        matched_pairs=[(row_a, row_b)],
        only_in_a_unique=[],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['id', 'name'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=['name'],
            exact_name_key_pairs=[ColumnPair('id', 'id')],
            exact_name_compare_pairs=[ColumnPair('name', 'name')],
            renamed_key_pairs=[],
            renamed_compare_pairs=[],
            unmapped_file_a_columns=[],
            unmapped_file_b_columns=[],
        ),
    )
    cell_diffs = [
        [
            CellDiff(
                column_a='name',
                column_b='name',
                row_key="id='1'",
                sheet='',
                value_a='@SUM(A1)',
                value_b='@SUM(A1)',
                are_equal=True,
            )
        ]
    ]
    result = ComparisonResult(
        matched=[
            RowDiff(
                row_key="id='1'",
                status=DiffStatus.MATCHED,
                cell_diffs=cell_diffs[0],
                index_a=0,
                index_b=0,
            )
        ],
        mismatched=[],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'matched_records': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 1},
    )
    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)
    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Unchanged Records']
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith('@'):
                assert cell.value == '@SUM(A1)', f"Expected '@SUM(A1)', got {cell.value!r}"
                assert cell.data_type == 's', f"Expected 's', got {cell.data_type!r}"


def test_write_matching_records_key_values_populated(tmp_path: Path):
    """Matching Records sheet writes actual key column values, not blanks."""
    row_a = _row({'id': '101', 'name': 'Alice', 'amount': '100'})
    row_b = _row({'id': '101', 'name': 'Alice', 'amount': '100'})
    key_pairs = [ColumnPair('id', 'id')]
    compare_pairs = [ColumnPair('amount', 'amount')]
    config = _config(key_pairs=key_pairs, compare_pairs=compare_pairs)

    match_result = MatchResult(
        matched_pairs=[(row_a, row_b)],
        only_in_a_unique=[],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['id', 'name', 'amount'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=['amount'],
        ),
    )

    result = ComparisonResult(
        matched=[
            RowDiff(
                row_key="id='101'",
                status=DiffStatus.MATCHED,
                cell_diffs=[
                    CellDiff(
                        column_a='amount',
                        column_b='amount',
                        row_key="id='101'",
                        sheet='',
                        value_a='100',
                        value_b='100',
                        are_equal=True,
                    )
                ],
                index_a=0,
                index_b=0,
                excel_row_a=2,
                excel_row_b=2,
                key_values_a=('101',),
                key_values_b=('101',),
            )
        ],
        mismatched=[],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'matched_records': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 1},
    )

    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)

    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Unchanged Records']

    headers = [cell.value for cell in ws[1]]
    assert 'File A – id' in headers
    assert 'File B – id' in headers

    data_rows = list(ws.iter_rows(min_row=2, values_only=True))
    assert len(data_rows) == 1
    # File A.id value should be '101', File B.id value should be '101'
    assert data_rows[0][0] == '101', f"Expected File A.id='101', got {data_rows[0][0]!r}"
    assert data_rows[0][1] == '101', f"Expected File B.id='101', got {data_rows[0][1]!r}"


def test_write_matching_records_composite_key_values(tmp_path: Path):
    """Composite keys preserve every individual key value."""
    row_a = _row({'dept': 'HR', 'emp_id': 'E001', 'salary': '5000'})
    row_b = _row({'dept': 'HR', 'emp_id': 'E001', 'salary': '5000'})
    key_pairs = [ColumnPair('dept', 'dept'), ColumnPair('emp_id', 'emp_id')]
    compare_pairs = [ColumnPair('salary', 'salary')]
    config = _config(key_pairs=key_pairs, compare_pairs=compare_pairs)

    match_result = MatchResult(
        matched_pairs=[(row_a, row_b)],
        only_in_a_unique=[],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['dept', 'emp_id', 'salary'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['dept', 'emp_id'],
            compare_columns=['salary'],
            exact_name_key_pairs=key_pairs,
            exact_name_compare_pairs=compare_pairs,
            renamed_key_pairs=[],
            renamed_compare_pairs=[],
            unmapped_file_a_columns=[],
            unmapped_file_b_columns=[],
        ),
    )

    result = ComparisonResult(
        matched=[
            RowDiff(
                row_key="dept='HR', emp_id='E001'",
                status=DiffStatus.MATCHED,
                cell_diffs=[
                    CellDiff(
                        column_a='salary',
                        column_b='salary',
                        row_key="dept='HR', emp_id='E001'",
                        sheet='',
                        value_a='5000',
                        value_b='5000',
                        are_equal=True,
                    )
                ],
                index_a=0,
                index_b=0,
                excel_row_a=2,
                excel_row_b=2,
                key_values_a=('HR', 'E001'),
                key_values_b=('HR', 'E001'),
            )
        ],
        mismatched=[],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'matched_records': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 1},
    )

    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)

    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Unchanged Records']

    headers = [cell.value for cell in ws[1]]
    assert 'File A – dept' in headers
    assert 'File B – dept' in headers
    assert 'File A – emp_id' in headers
    assert 'File B – emp_id' in headers

    data_rows = list(ws.iter_rows(min_row=2, values_only=True))
    assert len(data_rows) == 1
    assert data_rows[0][0] == 'HR'
    assert data_rows[0][1] == 'HR'
    assert data_rows[0][2] == 'E001'
    assert data_rows[0][3] == 'E001'


def test_write_matching_records_renamed_key_columns(tmp_path: Path):
    """Differently named key columns work correctly."""
    row_a = _row({'Site ID': 'S1', 'name': 'Alice', 'amount': '100'})
    row_b = _row({'Site Code': 'S1', 'name': 'Alice', 'amount': '100'})
    key_pairs = [ColumnPair('Site ID', 'Site Code')]
    compare_pairs = [ColumnPair('amount', 'amount')]
    config = _config(key_pairs=key_pairs, compare_pairs=compare_pairs)

    match_result = MatchResult(
        matched_pairs=[(row_a, row_b)],
        only_in_a_unique=[],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['name', 'amount'],
            only_in_a=['Site ID'],
            only_in_b=['Site Code'],
            key_columns=['Site ID'],
            compare_columns=['amount'],
            exact_name_key_pairs=[],
            exact_name_compare_pairs=compare_pairs,
            renamed_key_pairs=key_pairs,
            renamed_compare_pairs=[],
            unmapped_file_a_columns=[],
            unmapped_file_b_columns=[],
        ),
    )

    result = ComparisonResult(
        matched=[
            RowDiff(
                row_key="Site ID='S1', Site Code='S1'",
                status=DiffStatus.MATCHED,
                cell_diffs=[
                    CellDiff(
                        column_a='amount',
                        column_b='amount',
                        row_key="Site ID='S1', Site Code='S1'",
                        sheet='',
                        value_a='100',
                        value_b='100',
                        are_equal=True,
                    )
                ],
                index_a=0,
                index_b=0,
                excel_row_a=2,
                excel_row_b=2,
                key_values_a=('S1',),
                key_values_b=('S1',),
            )
        ],
        mismatched=[],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'matched_records': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 1},
    )

    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)

    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Unchanged Records']

    headers = [cell.value for cell in ws[1]]
    assert 'File A – Site ID' in headers
    assert 'File B – Site Code' in headers

    data_rows = list(ws.iter_rows(min_row=2, values_only=True))
    assert len(data_rows) == 1
    assert data_rows[0][0] == 'S1'
    assert data_rows[0][1] == 'S1'


def test_write_matching_records_backward_compatible_without_key_values(tmp_path: Path):
    """Existing RowDiff objects without key_values fields still write a valid report."""
    row_a = _row({'id': '1', 'amount': '100'})
    row_b = _row({'id': '1', 'amount': '100'})
    key_pairs = [ColumnPair('id', 'id')]
    compare_pairs = [ColumnPair('amount', 'amount')]
    config = _config(key_pairs=key_pairs, compare_pairs=compare_pairs)

    match_result = MatchResult(
        matched_pairs=[(row_a, row_b)],
        only_in_a_unique=[],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['id', 'amount'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=['amount'],
        ),
    )

    # RowDiff without key_values_a / key_values_b (backward compatible)
    result = ComparisonResult(
        matched=[
            RowDiff(
                row_key="id='1'",
                status=DiffStatus.MATCHED,
                cell_diffs=[
                    CellDiff(
                        column_a='amount',
                        column_b='amount',
                        row_key="id='1'",
                        sheet='',
                        value_a='100',
                        value_b='100',
                        are_equal=True,
                    )
                ],
                index_a=0,
                index_b=0,
            )
        ],
        mismatched=[],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'matched_records': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 1},
    )

    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)

    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Unchanged Records']

    data_rows = list(ws.iter_rows(min_row=2, values_only=True))
    assert len(data_rows) == 1
    # key_values_a/key_values_b default to empty tuple, so cells are blank
    assert data_rows[0][0] in ('', None)
    assert data_rows[0][1] in ('', None)


def _match_result_with_mismatch() -> tuple[MatchResult, ComparisonConfig]:
    """Helper: one mismatched record with differently named key columns."""
    row_a = _row({'Site ID': 'S1', 'name': 'Alice', 'amount': '100'})
    row_b = _row({'Site Code': 'S1', 'name': 'Alicia', 'amount': '100'})
    key_pairs = [ColumnPair('Site ID', 'Site Code')]
    compare_pairs = [ColumnPair('name', 'name'), ColumnPair('amount', 'amount')]
    config = _config(key_pairs=key_pairs, compare_pairs=compare_pairs)

    cell_diffs = [
        [
            CellDiff(
                column_a='name',
                column_b='name',
                row_key="Site ID='S1', Site Code='S1'",
                sheet='',
                value_a='Alice',
                value_b='Alicia',
                are_equal=False,
                excel_row_a=2,
                excel_row_b=2,
            ),
            CellDiff(
                column_a='amount',
                column_b='amount',
                row_key="Site ID='S1', Site Code='S1'",
                sheet='',
                value_a='100',
                value_b='100',
                are_equal=True,
                excel_row_a=2,
                excel_row_b=2,
            ),
        ]
    ]

    match_result = MatchResult(
        matched_pairs=[(row_a, row_b)],
        only_in_a_unique=[],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['name', 'amount'],
            only_in_a=['Site ID'],
            only_in_b=['Site Code'],
            key_columns=['Site ID'],
            compare_columns=['name', 'amount'],
            exact_name_key_pairs=[],
            exact_name_compare_pairs=compare_pairs,
            renamed_key_pairs=key_pairs,
            renamed_compare_pairs=[],
            unmapped_file_a_columns=[],
            unmapped_file_b_columns=[],
        ),
    )

    result = ComparisonResult(
        matched=[],
        mismatched=[
            RowDiff(
                row_key="Site ID='S1', Site Code='S1'",
                status=DiffStatus.MISMATCHED,
                cell_diffs=cell_diffs[0],
                index_a=0,
                index_b=0,
                excel_row_a=2,
                excel_row_b=2,
                key_values_a=('S1',),
                key_values_b=('S1',),
            )
        ],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[
            DataQualityIssue(
                issue_type='DUPLICATE_KEY',
                source_file='File A',
                original_index=0,
                excel_row_number=2,
                key_display="Site ID='S1'",
                all_column_values={'Site ID': 'S1', 'name': 'Alice'},
            )
        ],
        schema_diff=match_result.schema_diff,
        summary={
            'total_rows_file_a': 1,
            'total_rows_file_b': 1,
            'matched_records': 0,
            'mismatched_records': 1,
            'matched_no_fields_compared': 0,
            'only_in_file_a': 0,
            'only_in_file_b': 0,
            'duplicate_key_records': 1,
            'missing_key_records': 0,
        },
    )
    return result, config


def test_write_records_with_changes_client_friendly_headings(tmp_path: Path):
    """Records With Changes sheet uses Record ID / field / value headings."""
    result, config = _match_result_with_mismatch()
    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)

    workbook = openpyxl.load_workbook(output_path)
    assert 'Records With Changes' in workbook.sheetnames
    ws = workbook['Records With Changes']
    headers = [cell.value for cell in ws[1]]
    assert 'Record ID' not in headers
    assert 'File A – Site ID' in headers
    assert 'File B – Site Code' in headers
    assert 'Column Name (File A)' in headers
    assert 'Previous Value' in headers
    assert 'Column Name (File B)' in headers
    assert 'Updated Value' in headers
    assert 'Excel Row (File A)' in headers
    assert 'Excel Row (File B)' in headers

    data_rows = list(ws.iter_rows(min_row=2, values_only=True))
    assert len(data_rows) == 1
    assert data_rows[0][0] == 'S1'
    assert data_rows[0][1] == 'S1'
    assert data_rows[0][2] == 'name'
    assert data_rows[0][3] == 'Alice'
    assert data_rows[0][4] == 'name'
    assert data_rows[0][5] == 'Alicia'
    assert data_rows[0][6] == 2
    assert data_rows[0][7] == 2


def test_write_unchanged_records_preserves_key_values(tmp_path: Path):
    """Unchanged Records sheet writes actual key column values under file headers."""
    result, config = _match_result_with_mismatch()
    # Convert the mismatched record into a matched one to land on Unchanged Records.
    matched = RowDiff(
        row_key="Site ID='S1', Site Code='S1'",
        status=DiffStatus.MATCHED,
        cell_diffs=[],
        index_a=0,
        index_b=0,
        excel_row_a=2,
        excel_row_b=2,
        key_values_a=('S1',),
        key_values_b=('S1',),
    )
    result = ComparisonResult(
        matched=[matched],
        mismatched=[],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=result.schema_diff,
        summary=result.summary,
    )
    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)

    workbook = openpyxl.load_workbook(output_path)
    assert 'Unchanged Records' in workbook.sheetnames
    ws = workbook['Unchanged Records']
    headers = [cell.value for cell in ws[1]]
    assert 'File A – Site ID' in headers
    assert 'File B – Site Code' in headers
    data_rows = list(ws.iter_rows(min_row=2, values_only=True))
    assert len(data_rows) == 1
    assert data_rows[0][0] == 'S1'
    assert data_rows[0][1] == 'S1'


def test_write_formula_injection_protected_on_mismatched_sheet(tmp_path: Path):
    """Strings starting with '=' are literal text on Records With Changes."""
    row_a = _row({'id': '1', 'name': '=SUM(A1:A2)'})
    row_b = _row({'id': '1', 'name': '=SUM(A1:A2)'})
    key_pairs = [ColumnPair('id', 'id')]
    compare_pairs = [ColumnPair('name', 'name')]
    config = _config(key_pairs=key_pairs, compare_pairs=compare_pairs)

    cell_diffs = [
        [
            CellDiff(
                column_a='name',
                column_b='name',
                row_key="id='1'",
                sheet='',
                value_a='=SUM(A1:A2)',
                value_b='=SUM(A1:A2)',
                are_equal=False,
                excel_row_a=2,
                excel_row_b=2,
            )
        ]
    ]
    match_result = MatchResult(
        matched_pairs=[(row_a, row_b)],
        only_in_a_unique=[],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['id', 'name'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=['name'],
        ),
    )
    result = ComparisonResult(
        matched=[],
        mismatched=[
            RowDiff(
                row_key="id='1'",
                status=DiffStatus.MISMATCHED,
                cell_diffs=cell_diffs[0],
                index_a=0,
                index_b=0,
                excel_row_a=2,
                excel_row_b=2,
                key_values_a=('1',),
                key_values_b=('1',),
            )
        ],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'mismatched_records': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 1},
    )

    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)

    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Records With Changes']
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith('='):
                assert cell.value == '=SUM(A1:A2)', f'Expected literal text, got {cell.value!r}'
                assert cell.data_type == 's', f"Expected 's', got {cell.data_type!r}"


def test_write_only_in_sheets_use_record_id_heading(tmp_path: Path):
    """Missing From File B / New Records in File B use Record ID heading."""
    row_a = _row({'id': '9', 'name': 'Orphan'})
    config = _config(key_pairs=[ColumnPair('id', 'id')], compare_pairs=[ColumnPair('name', 'name')])
    match_result = MatchResult(
        matched_pairs=[],
        only_in_a_unique=[row_a],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['id', 'name'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=['name'],
        ),
    )
    result = ComparisonResult(
        matched=[],
        mismatched=[],
        only_in_a=[
            RowDiff(
                row_key="id='9'",
                status=DiffStatus.ONLY_IN_A,
                cell_diffs=[],
                index_a=0,
                excel_row_a=5,
                key_values_a=('9',),
            )
        ],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'only_in_file_a': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 0},
    )

    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config, df_a=pd.DataFrame([dict(row_a)]))

    workbook = openpyxl.load_workbook(output_path)
    assert 'Missing From File B' in workbook.sheetnames
    ws = workbook['Missing From File B']
    headers = [cell.value for cell in ws[1]]
    assert 'Record ID' not in headers
    assert 'id' in headers
    assert 'name' in headers
    assert 'Record Status' not in headers
    assert 'Excel Row (File A)' in headers
    data_rows = list(ws.iter_rows(min_row=2, values_only=True))
    assert len(data_rows) == 1
    assert data_rows[0][0] == '9'
    assert data_rows[0][1] == 'Orphan'
    assert data_rows[0][2] == 5


def test_write_comparison_overview_plain_language_labels(tmp_path: Path):
    """Comparison Overview sheet uses plain-language metric labels."""
    result, config = _match_result_with_mismatch()
    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)

    workbook = openpyxl.load_workbook(output_path)
    assert 'Comparison Overview' in workbook.sheetnames
    ws = workbook['Comparison Overview']
    labels = [cell.value for cell in ws['A']]
    assert 'Records in File A' in labels
    assert 'Records with no changes' in labels
    assert 'Records with changes' in labels
    assert 'Records matched but not compared' in labels
    assert 'Duplicate identification records' in labels
    assert 'Records with missing identification values' in labels
    assert 'Records missing from File B' in labels
    assert 'New records in File B' in labels


# ---------------------------------------------------------------------------
# Per-file key column regression tests
# ---------------------------------------------------------------------------


def _composite_match_result() -> tuple[MatchResult, ComparisonConfig]:
    """Helper: composite key matched record."""
    row_a = _row(
        {'SR NO': '51', 'snapshot_date': '2026-01-01', 'warehouse': 'DC-Atlanta', 'sku': 'SKU-1051'}
    )
    row_b = _row(
        {'SR NO': '51', 'snapshot_date': '2026-01-01', 'warehouse': 'DC-Atlanta', 'sku': 'SKU-1051'}
    )
    key_pairs = [
        ColumnPair('SR NO', 'SR NO'),
        ColumnPair('snapshot_date', 'snapshot_date'),
        ColumnPair('warehouse', 'warehouse'),
        ColumnPair('sku', 'sku'),
    ]
    compare_pairs = []
    config = _config(key_pairs=key_pairs, compare_pairs=compare_pairs)

    match_result = MatchResult(
        matched_pairs=[(row_a, row_b)],
        only_in_a_unique=[],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['SR NO', 'snapshot_date', 'warehouse', 'sku'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['SR NO', 'snapshot_date', 'warehouse', 'sku'],
            compare_columns=[],
        ),
    )
    result = ComparisonResult(
        matched=[
            RowDiff(
                row_key="SR NO='51', snapshot_date='2026-01-01', warehouse='DC-Atlanta', sku='SKU-1051'",
                status=DiffStatus.MATCHED,
                cell_diffs=[],
                index_a=0,
                index_b=0,
                excel_row_a=2,
                excel_row_b=2,
                key_values_a=('51', '2026-01-01', 'DC-Atlanta', 'SKU-1051'),
                key_values_b=('51', '2026-01-01', 'DC-Atlanta', 'SKU-1051'),
            )
        ],
        mismatched=[],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'matched_records': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 1},
    )
    return result, config


def test_writer_unchanged_records_per_file_key_columns(tmp_path: Path):
    """Unchanged Records sheet shows both files' key columns separately."""
    result, config = _composite_match_result()
    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)

    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Unchanged Records']
    headers = [cell.value for cell in ws[1]]

    # Both File A and File B key columns present, in mapping order
    assert 'File A – SR NO' in headers
    assert 'File B – SR NO' in headers
    assert 'File A – snapshot_date' in headers
    assert 'File B – snapshot_date' in headers
    assert 'File A – warehouse' in headers
    assert 'File B – warehouse' in headers
    assert 'File A – sku' in headers
    assert 'File B – sku' in headers
    # No combined Record ID / Row Key column
    assert 'Record ID' not in headers
    assert 'Row Key' not in headers

    data_rows = list(ws.iter_rows(min_row=2, values_only=True))
    assert len(data_rows) == 1
    # File A values | File B values (interleaved by key pair order)
    assert data_rows[0][0] == '51'  # File A.SR NO
    assert data_rows[0][1] == '51'  # File B.SR NO
    assert data_rows[0][2] == '2026-01-01'  # File A.snapshot_date
    assert data_rows[0][3] == '2026-01-01'  # File B.snapshot_date
    assert data_rows[0][4] == 'DC-Atlanta'  # File A.warehouse
    assert data_rows[0][5] == 'DC-Atlanta'  # File B.warehouse
    assert data_rows[0][6] == 'SKU-1051'  # File A.sku
    assert data_rows[0][7] == 'SKU-1051'  # File B.sku


def test_writer_only_in_a_blank_file_b_keys(tmp_path: Path):
    """Only-in-A records have File A keys populated, File B keys blank."""
    row_a = _row({'id': '9', 'name': 'Orphan'})
    config = _config(key_pairs=[ColumnPair('id', 'id')], compare_pairs=[ColumnPair('name', 'name')])
    match_result = MatchResult(
        matched_pairs=[],
        only_in_a_unique=[row_a],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['id', 'name'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=['name'],
        ),
    )
    result = ComparisonResult(
        matched=[],
        mismatched=[],
        only_in_a=[
            RowDiff(
                row_key="id='9'",
                status=DiffStatus.ONLY_IN_A,
                cell_diffs=[],
                index_a=0,
                excel_row_a=5,
                key_values_a=('9',),
                key_values_b=(),
            )
        ],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'only_in_file_a': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 0},
    )

    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config, df_a=pd.DataFrame([dict(row_a)]))

    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Missing From File B']
    headers = [cell.value for cell in ws[1]]

    assert 'id' in headers
    assert 'name' in headers
    assert 'Record ID' not in headers
    assert 'Record Status' not in headers

    data_rows = list(ws.iter_rows(min_row=2, values_only=True))
    assert len(data_rows) == 1
    assert data_rows[0][0] == '9'
    assert data_rows[0][1] == 'Orphan'


def test_writer_only_in_b_blank_file_a_keys(tmp_path: Path):
    """Only-in-B records have File B keys populated, File A keys blank."""
    row_b = _row({'id': '9', 'name': 'Newbie'})
    config = _config(key_pairs=[ColumnPair('id', 'id')], compare_pairs=[ColumnPair('name', 'name')])
    match_result = MatchResult(
        matched_pairs=[],
        only_in_a_unique=[],
        only_in_b_unique=[row_b],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['id', 'name'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=['name'],
        ),
    )
    result = ComparisonResult(
        matched=[],
        mismatched=[],
        only_in_a=[],
        only_in_b=[
            RowDiff(
                row_key="id='9'",
                status=DiffStatus.ONLY_IN_B,
                cell_diffs=[],
                index_b=0,
                excel_row_b=5,
                key_values_a=(),
                key_values_b=('9',),
            )
        ],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'only_in_file_b': 1, 'total_rows_file_a': 0, 'total_rows_file_b': 1},
    )

    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config, df_b=pd.DataFrame([dict(row_b)]))

    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['New Records in File B']
    headers = [cell.value for cell in ws[1]]

    assert 'id' in headers
    assert 'name' in headers
    assert 'Record ID' not in headers
    assert 'Record Status' not in headers

    data_rows = list(ws.iter_rows(min_row=2, values_only=True))
    assert len(data_rows) == 1
    assert data_rows[0][0] == '9'
    assert data_rows[0][1] == 'Newbie'


def test_writer_data_quality_preserves_per_file_keys(tmp_path: Path):
    """Records Needing Attention shows per-file key values, not parsed key_display."""
    dup_rec = DuplicateSourceRecord(
        source_file='A',
        original_index=5,
        excel_row_number=7,
        key_values=('1',),
        key_display="id='1'",
        group_occurrence_count=2,
        all_column_values={'id': '1', 'name': 'Alice'},
    )
    config = _config(key_pairs=[ColumnPair('id', 'id')], compare_columns=['name'])
    match_result = MatchResult(
        matched_pairs=[],
        only_in_a_unique=[],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[dup_rec],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['id', 'name'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=['name'],
        ),
    )
    result = ComparisonResult(
        matched=[],
        mismatched=[],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[
            DataQualityIssue(
                issue_type='DUPLICATE_KEY',
                source_file='A',
                original_index=5,
                excel_row_number=7,
                key_display="id='1'",
                all_column_values={'id': '1', 'name': 'Alice'},
                key_values_a=('1',),
                key_values_b=(),
            )
        ],
        schema_diff=match_result.schema_diff,
        summary={'duplicate_key_records': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 0},
    )

    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)

    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Records Needing Attention']
    headers = [cell.value for cell in ws[1]]

    assert 'File A – id' in headers
    assert 'File B – id' in headers
    assert 'Record ID' not in headers

    data_rows = list(ws.iter_rows(min_row=2, values_only=True))
    assert len(data_rows) == 1
    assert data_rows[0][3] == '1'  # File A.id (after Issue Type, Source File, Excel Row)
    assert data_rows[0][4] in ('', None)  # File B.id blank


def test_writer_no_combined_record_id_column_anywhere(tmp_path: Path):
    """No Record ID / Row Key column appears in any record-level sheet."""
    result, config = _composite_match_result()
    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)

    workbook = openpyxl.load_workbook(output_path)

    for sheet_name in [
        'Unchanged Records',
        'Records With Changes',
        'Missing From File B',
        'New Records in File B',
        'Records Needing Attention',
    ]:
        if sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]
            headers = [cell.value for cell in ws[1]]
            assert 'Record ID' not in headers, f'Record ID found in {sheet_name}'
            assert 'Row Key' not in headers, f'Row Key found in {sheet_name}'


def test_writer_formula_injection_on_key_columns(tmp_path: Path):
    """Formula-injection protection applies to key value cells."""
    row_a = _row({'id': '=1+1', 'name': 'Alice'})
    row_b = _row({'id': '=1+1', 'name': 'Alice'})
    key_pairs = [ColumnPair('id', 'id')]
    compare_pairs = [ColumnPair('name', 'name')]
    config = _config(key_pairs=key_pairs, compare_pairs=compare_pairs)

    match_result = MatchResult(
        matched_pairs=[(row_a, row_b)],
        only_in_a_unique=[],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['id', 'name'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=['name'],
        ),
    )
    result = ComparisonResult(
        matched=[
            RowDiff(
                row_key="id='=1+1'",
                status=DiffStatus.MATCHED,
                cell_diffs=[],
                index_a=0,
                index_b=0,
                excel_row_a=2,
                excel_row_b=2,
                key_values_a=('=1+1',),
                key_values_b=('=1+1',),
            )
        ],
        mismatched=[],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'matched_records': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 1},
    )

    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)

    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Unchanged Records']
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith('='):
                assert cell.value == '=1+1', f'Expected literal text, got {cell.value!r}'
                assert cell.data_type == 's', f"Expected 's', got {cell.data_type!r}"


# ---------------------------------------------------------------------------
# Column heading regression tests
# ---------------------------------------------------------------------------


def _basic_config() -> ComparisonConfig:
    key_pairs = [ColumnPair('SR NO', 'SR'), ColumnPair('snapshot_date', 'snapshot_date')]
    compare_pairs = [ColumnPair('warehouse', 'warehouse')]
    return _config(key_pairs=key_pairs, compare_pairs=compare_pairs)


def _matched_result_with_df(config: ComparisonConfig) -> ComparisonResult:
    row_a = _row(
        {'SR NO': '101', 'snapshot_date': '2026-01-01', 'warehouse': 'Jaipur', 'quantity': '50'}
    )
    row_b = _row(
        {'SR': '101', 'snapshot_date': '2026-01-01', 'warehouse': 'Jaipur', 'quantity': '50'}
    )
    match_result = MatchResult(
        matched_pairs=[(row_a, row_b)],
        only_in_a_unique=[],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['SR NO', 'SR', 'snapshot_date', 'warehouse', 'quantity'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['SR NO', 'SR'],
            compare_columns=['warehouse', 'quantity'],
        ),
    )
    return ComparisonResult(
        matched=[
            RowDiff(
                row_key="SR NO='101', SR='101'",
                status=DiffStatus.MATCHED,
                cell_diffs=[],
                index_a=0,
                index_b=0,
                excel_row_a=2,
                excel_row_b=2,
                key_values_a=('101', '2026-01-01'),
                key_values_b=('101', '2026-01-01'),
            )
        ],
        mismatched=[],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'matched_records': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 1},
    )


def test_writer_column_heading_file_a_sr_no(tmp_path: Path):
    """File A.SR NO becomes File A – SR NO."""
    config = _basic_config()
    result = _matched_result_with_df(config)
    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)
    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Unchanged Records']
    headers = [cell.value for cell in ws[1]]
    assert 'File A – SR NO' in headers


def test_writer_column_heading_file_a_snapshot_date(tmp_path: Path):
    """File A.snapshot_date becomes File A – snapshot_date."""
    config = _basic_config()
    result = _matched_result_with_df(config)
    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)
    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Unchanged Records']
    headers = [cell.value for cell in ws[1]]
    assert 'File A – snapshot_date' in headers


def test_writer_column_heading_file_b_sr(tmp_path: Path):
    """File B.SR becomes File B – SR."""
    config = _basic_config()
    result = _matched_result_with_df(config)
    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)
    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Unchanged Records']
    headers = [cell.value for cell in ws[1]]
    assert 'File B – SR' in headers


def test_writer_column_heading_file_b_snapshot_date(tmp_path: Path):
    """File B.snapshot_date becomes File B – snapshot_date."""
    config = _basic_config()
    result = _matched_result_with_df(config)
    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)
    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Unchanged Records']
    headers = [cell.value for cell in ws[1]]
    assert 'File B – snapshot_date' in headers


def test_writer_column_heading_dynamic_arbitrary(tmp_path: Path):
    """Arbitrary uploaded column names are handled dynamically."""
    key_pairs = [ColumnPair('employee_id', 'emp_id'), ColumnPair('warehouse_code', 'wh_code')]
    compare_pairs = [ColumnPair('amount', 'amt')]
    config = _config(key_pairs=key_pairs, compare_pairs=compare_pairs)
    row_a = _row({'employee_id': 'E1', 'warehouse_code': 'WH1', 'amount': '100'})
    row_b = _row({'emp_id': 'E1', 'wh_code': 'WH1', 'amt': '100'})
    match_result = MatchResult(
        matched_pairs=[(row_a, row_b)],
        only_in_a_unique=[],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['employee_id', 'warehouse_code', 'amount'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['employee_id', 'warehouse_code'],
            compare_columns=['amount'],
        ),
    )
    result = ComparisonResult(
        matched=[
            RowDiff(
                row_key="employee_id='E1', warehouse_code='WH1'",
                status=DiffStatus.MATCHED,
                cell_diffs=[],
                index_a=0,
                index_b=0,
                excel_row_a=2,
                excel_row_b=2,
                key_values_a=('E1', 'WH1'),
                key_values_b=('E1', 'WH1'),
            )
        ],
        mismatched=[],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'matched_records': 1},
    )
    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)
    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Unchanged Records']
    headers = [cell.value for cell in ws[1]]
    assert 'File A – employee_id' in headers
    assert 'File B – emp_id' in headers
    assert 'File A – warehouse_code' in headers
    assert 'File B – wh_code' in headers


def test_writer_column_heading_original_preserved(tmp_path: Path):
    """Original column names remain unchanged after prefix."""
    config = _basic_config()
    result = _matched_result_with_df(config)
    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)
    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Unchanged Records']
    headers = [cell.value for cell in ws[1]]
    # Original names preserved exactly (no title case, no acronym expansion)
    assert 'File A – SR NO' in headers
    assert 'File A – snapshot_date' in headers
    assert 'File B – SR' in headers
    assert 'File B – snapshot_date' in headers
    assert 'Snapshot Date' not in headers
    assert 'Serial Number' not in headers


def test_writer_all_record_sheets_both_file_headings(tmp_path: Path):
    """Every record-level sheet has both File A and File B key headings."""
    config = _basic_config()
    result = _matched_result_with_df(config)

    # Missing From File B
    row_a = _row(
        {'SR NO': '9', 'snapshot_date': '2026-01-01', 'warehouse': 'Jaipur', 'quantity': '50'}
    )
    only_in_a_result = ComparisonResult(
        matched=[],
        mismatched=[],
        only_in_a=[
            RowDiff(
                row_key="SR NO='9'",
                status=DiffStatus.ONLY_IN_A,
                cell_diffs=[],
                index_a=0,
                excel_row_a=5,
                key_values_a=('9', '2026-01-01'),
                key_values_b=(),
            )
        ],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=result.schema_diff,
        summary={'only_in_file_a': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 0},
    )
    output_path = tmp_path / 'report.xlsx'
    write_report(only_in_a_result, str(output_path), config, df_a=pd.DataFrame([dict(row_a)]))
    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Missing From File B']
    headers = [cell.value for cell in ws[1]]
    assert 'SR NO' in headers
    assert 'snapshot_date' in headers
    assert 'warehouse' in headers
    assert 'quantity' in headers
    assert 'Record ID' not in headers
    assert 'Record Status' not in headers
    assert 'Excel Row (File A)' in headers


def test_writer_excel_row_headings_consistent(tmp_path: Path):
    """Excel Row headings are consistent across sheets."""
    config = _basic_config()
    result = _matched_result_with_df(config)
    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)
    workbook = openpyxl.load_workbook(output_path)

    ws = workbook['Unchanged Records']
    headers = [cell.value for cell in ws[1]]
    assert 'Excel Row (File A)' in headers
    assert 'Excel Row (File B)' in headers

    ws = workbook['Records Needing Attention']
    headers = [cell.value for cell in ws[1]]
    assert 'Excel Row' in headers


def test_writer_six_sheet_names_unchanged(tmp_path: Path):
    """All six report sheet names remain unchanged."""
    config = _basic_config()
    result = _matched_result_with_df(config)
    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)
    workbook = openpyxl.load_workbook(output_path)
    assert workbook.sheetnames == list(REPORT_SHEETS)
    assert len(workbook.sheetnames) == 6


def test_writer_formula_injection_still_protected(tmp_path: Path):
    """Formula-injection protection remains effective on key and data columns."""
    config = _basic_config()
    row_a = _row(
        {'SR NO': '=1+1', 'snapshot_date': '2026-01-01', 'warehouse': 'Jaipur', 'quantity': '50'}
    )
    row_b = _row(
        {'SR': '=1+1', 'snapshot_date': '2026-01-01', 'warehouse': 'Jaipur', 'quantity': '50'}
    )
    match_result = MatchResult(
        matched_pairs=[(row_a, row_b)],
        only_in_a_unique=[],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['SR NO', 'SR', 'snapshot_date', 'warehouse', 'quantity'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['SR NO', 'SR'],
            compare_columns=['warehouse', 'quantity'],
        ),
    )
    result = ComparisonResult(
        matched=[
            RowDiff(
                row_key="SR NO='=1+1'",
                status=DiffStatus.MATCHED,
                cell_diffs=[],
                index_a=0,
                index_b=0,
                excel_row_a=2,
                excel_row_b=2,
                key_values_a=('=1+1', '2026-01-01'),
                key_values_b=('=1+1', '2026-01-01'),
            )
        ],
        mismatched=[],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'matched_records': 1},
    )
    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config)
    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Unchanged Records']
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith('='):
                assert cell.value == '=1+1', f'Expected literal text, got {cell.value!r}'
                assert cell.data_type == 's', f"Expected 's', got {cell.data_type!r}"


# ---------------------------------------------------------------------------
# Complete-record regression tests
# ---------------------------------------------------------------------------


def test_writer_missing_from_file_b_complete_record(tmp_path: Path):
    """Missing From File B sheet contains complete original File A record."""
    key_pairs = [ColumnPair('id', 'id')]
    compare_pairs = [ColumnPair('name', 'name')]
    config = _config(key_pairs=key_pairs, compare_pairs=compare_pairs)
    row_a = _row({'id': '9', 'name': 'Orphan', 'extra_col': 'unmapped', 'amount': '999'})
    df_a = pd.DataFrame([dict(row_a)])

    match_result = MatchResult(
        matched_pairs=[],
        only_in_a_unique=[row_a],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['id', 'name'],
            only_in_a=['extra_col'],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=['name'],
        ),
    )
    result = ComparisonResult(
        matched=[],
        mismatched=[],
        only_in_a=[
            RowDiff(
                row_key="id='9'",
                status=DiffStatus.ONLY_IN_A,
                cell_diffs=[],
                index_a=0,
                excel_row_a=5,
                key_values_a=('9',),
                key_values_b=(),
            )
        ],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'only_in_file_a': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 0},
    )

    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config, df_a=df_a, df_b=None)

    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Missing From File B']
    headers = [cell.value for cell in ws[1]]

    # Complete File A record in original column order
    assert 'id' in headers
    assert 'name' in headers
    assert 'extra_col' in headers
    assert 'amount' in headers
    assert 'Record Status' not in headers
    assert 'Excel Row (File A)' in headers

    data_rows = list(ws.iter_rows(min_row=2, values_only=True))
    assert len(data_rows) == 1
    # Complete File A record values
    col_idx = {h: i for i, h in enumerate(headers)}
    assert data_rows[0][col_idx['id']] == '9'
    assert data_rows[0][col_idx['name']] == 'Orphan'
    assert data_rows[0][col_idx['extra_col']] == 'unmapped'
    assert data_rows[0][col_idx['amount']] == '999'
    assert data_rows[0][col_idx['Excel Row (File A)']] == 5


def test_writer_new_in_file_b_complete_record(tmp_path: Path):
    """New Records in File B sheet contains complete original File B record."""
    key_pairs = [ColumnPair('id', 'id')]
    compare_pairs = [ColumnPair('name', 'name')]
    config = _config(key_pairs=key_pairs, compare_pairs=compare_pairs)
    row_b = _row({'id': '9', 'name': 'Newbie', 'extra_col_b': 'unmapped_b', 'amount': '888'})
    df_b = pd.DataFrame([dict(row_b)])

    match_result = MatchResult(
        matched_pairs=[],
        only_in_a_unique=[],
        only_in_b_unique=[row_b],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['id', 'name'],
            only_in_a=[],
            only_in_b=['extra_col_b'],
            key_columns=['id'],
            compare_columns=['name'],
        ),
    )
    result = ComparisonResult(
        matched=[],
        mismatched=[],
        only_in_a=[],
        only_in_b=[
            RowDiff(
                row_key="id='9'",
                status=DiffStatus.ONLY_IN_B,
                cell_diffs=[],
                index_b=0,
                excel_row_b=5,
                key_values_a=(),
                key_values_b=('9',),
            )
        ],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'only_in_file_b': 1, 'total_rows_file_a': 0, 'total_rows_file_b': 1},
    )

    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config, df_a=None, df_b=df_b)

    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['New Records in File B']
    headers = [cell.value for cell in ws[1]]

    # Complete File B record in original column order
    assert 'id' in headers
    assert 'name' in headers
    assert 'extra_col_b' in headers
    assert 'amount' in headers
    assert 'Record Status' not in headers
    assert 'Excel Row (File B)' in headers

    data_rows = list(ws.iter_rows(min_row=2, values_only=True))
    assert len(data_rows) == 1
    # Complete File B record values
    col_idx = {h: i for i, h in enumerate(headers)}
    assert data_rows[0][col_idx['id']] == '9'
    assert data_rows[0][col_idx['name']] == 'Newbie'
    assert data_rows[0][col_idx['extra_col_b']] == 'unmapped_b'
    assert data_rows[0][col_idx['amount']] == '888'
    assert data_rows[0][col_idx['Excel Row (File B)']] == 5


def test_writer_missing_and_new_no_duplicate_key_columns(tmp_path: Path):
    """Key columns are not duplicated in remaining source columns."""
    key_pairs = [ColumnPair('id', 'id'), ColumnPair('code', 'code')]
    compare_pairs = [ColumnPair('name', 'name')]
    config = _config(key_pairs=key_pairs, compare_pairs=compare_pairs)
    row_a = _row({'id': '1', 'code': 'X', 'name': 'A', 'id_dup': 'ignore'})
    df_a = pd.DataFrame([dict(row_a)])

    match_result = MatchResult(
        matched_pairs=[],
        only_in_a_unique=[row_a],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['id', 'code', 'name'],
            only_in_a=['id_dup'],
            only_in_b=[],
            key_columns=['id', 'code'],
            compare_columns=['name'],
        ),
    )
    result = ComparisonResult(
        matched=[],
        mismatched=[],
        only_in_a=[
            RowDiff(
                row_key="id='1', code='X'",
                status=DiffStatus.ONLY_IN_A,
                cell_diffs=[],
                index_a=0,
                excel_row_a=3,
                key_values_a=('1', 'X'),
                key_values_b=(),
            )
        ],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'only_in_file_a': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 0},
    )

    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config, df_a=df_a, df_b=None)

    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Missing From File B']
    headers = [cell.value for cell in ws[1]]

    # All original columns appear exactly once
    assert headers.count('id') == 1
    assert headers.count('code') == 1
    assert headers.count('name') == 1
    assert headers.count('id_dup') == 1
    assert 'Excel Row (File A)' in headers


# ---------------------------------------------------------------------------
# Native data type preservation in report sheets (Issue 1)
# ---------------------------------------------------------------------------


def test_writer_missing_from_file_b_preserves_native_types(tmp_path: Path):
    """Missing From File B sheet preserves native int/float/date values."""
    key_pairs = [ColumnPair('id', 'id')]
    compare_pairs = [ColumnPair('name', 'name')]
    config = _config(key_pairs=key_pairs, compare_pairs=compare_pairs)
    row_a = _row({'id': 42, 'name': 'Orphan', 'amount': 99.5, 'd': date(2024, 1, 15)})
    df_a = pd.DataFrame([dict(row_a)])

    match_result = MatchResult(
        matched_pairs=[],
        only_in_a_unique=[row_a],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['id', 'name', 'amount', 'd'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=['name'],
        ),
    )
    result = ComparisonResult(
        matched=[],
        mismatched=[],
        only_in_a=[
            RowDiff(
                row_key="id='42'",
                status=DiffStatus.ONLY_IN_A,
                cell_diffs=[],
                index_a=0,
                excel_row_a=5,
                key_values_a=('42',),
                key_values_b=(),
            )
        ],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'only_in_file_a': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 0},
    )

    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config, df_a=df_a, df_b=None)

    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Missing From File B']
    headers = [cell.value for cell in ws[1]]
    col_idx = {h: i for i, h in enumerate(headers)}

    data_rows = list(ws.iter_rows(min_row=2, values_only=True))
    assert len(data_rows) == 1
    row = data_rows[0]
    # Native types preserved in the report
    assert row[col_idx['id']] == 42
    assert isinstance(row[col_idx['id']], int)
    assert row[col_idx['name']] == 'Orphan'
    assert row[col_idx['amount']] == 99.5
    assert isinstance(row[col_idx['amount']], float)
    # Date is written as datetime; compare the date part.
    assert row[col_idx['d']] == datetime(2024, 1, 15, 0, 0)


def test_writer_new_in_file_b_preserves_native_types(tmp_path: Path):
    """New Records in File B sheet preserves native int/float/date values."""
    key_pairs = [ColumnPair('id', 'id')]
    compare_pairs = [ColumnPair('name', 'name')]
    config = _config(key_pairs=key_pairs, compare_pairs=compare_pairs)
    row_b = _row({'id': 42, 'name': 'Newbie', 'amount': 99.5, 'd': date(2024, 1, 15)})
    df_b = pd.DataFrame([dict(row_b)])

    match_result = MatchResult(
        matched_pairs=[],
        only_in_a_unique=[],
        only_in_b_unique=[row_b],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['id', 'name', 'amount', 'd'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=['name'],
        ),
    )
    result = ComparisonResult(
        matched=[],
        mismatched=[],
        only_in_a=[],
        only_in_b=[
            RowDiff(
                row_key="id='42'",
                status=DiffStatus.ONLY_IN_B,
                cell_diffs=[],
                index_b=0,
                excel_row_b=5,
                key_values_a=(),
                key_values_b=('42',),
            )
        ],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'only_in_file_b': 1, 'total_rows_file_a': 0, 'total_rows_file_b': 1},
    )

    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config, df_a=None, df_b=df_b)

    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['New Records in File B']
    headers = [cell.value for cell in ws[1]]
    col_idx = {h: i for i, h in enumerate(headers)}

    data_rows = list(ws.iter_rows(min_row=2, values_only=True))
    assert len(data_rows) == 1
    row = data_rows[0]
    # Native types preserved in the report
    assert row[col_idx['id']] == 42
    assert isinstance(row[col_idx['id']], int)
    assert row[col_idx['name']] == 'Newbie'
    assert row[col_idx['amount']] == 99.5
    assert isinstance(row[col_idx['amount']], float)
    # Date is written as datetime; compare the date part.
    assert row[col_idx['d']] == datetime(2024, 1, 15, 0, 0)


def test_writer_formula_injection_with_numeric_values(tmp_path: Path):
    """Numeric cells are written as numbers (not text), strings starting with = remain text."""
    key_pairs = [ColumnPair('id', 'id')]
    compare_pairs = [ColumnPair('name', 'name')]
    config = _config(key_pairs=key_pairs, compare_pairs=compare_pairs)
    row_a = _row({'id': 42, 'name': '=1+1', 'amount': 100.5})
    df_a = pd.DataFrame([dict(row_a)])

    match_result = MatchResult(
        matched_pairs=[],
        only_in_a_unique=[row_a],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[],
        schema_diff=SchemaDiff(
            common_columns=['id', 'name', 'amount'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=['name'],
        ),
    )
    result = ComparisonResult(
        matched=[],
        mismatched=[],
        only_in_a=[
            RowDiff(
                row_key="id='42'",
                status=DiffStatus.ONLY_IN_A,
                cell_diffs=[],
                index_a=0,
                excel_row_a=5,
                key_values_a=('42',),
                key_values_b=(),
            )
        ],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=match_result.schema_diff,
        summary={'only_in_file_a': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 0},
    )

    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config, df_a=df_a, df_b=None)

    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Missing From File B']
    headers = [cell.value for cell in ws[1]]
    col_idx = {h: i for i, h in enumerate(headers)}

    data_rows = list(ws.iter_rows(min_row=2, values_only=True))
    assert len(data_rows) == 1
    row = data_rows[0]
    # Numeric values are numbers, not text
    assert row[col_idx['id']] == 42
    assert isinstance(row[col_idx['id']], int)
    assert row[col_idx['amount']] == 100.5
    assert isinstance(row[col_idx['amount']], float)
    # Formula-injection string remains literal text with data_type='s'
    assert row[col_idx['name']] == '=1+1'
    cell = ws.cell(row=2, column=col_idx['name'] + 1)
    assert cell.data_type == 's'


# ---------------------------------------------------------------------------
# Workbook resource cleanup
# ---------------------------------------------------------------------------


def test_write_report_closes_workbook(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    """write_report closes its workbook after a successful save."""
    instances: list = []

    class TrackingWorkbook(openpyxl.Workbook):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self.closed = False
            instances.append(self)

        def close(self) -> None:
            self.closed = True
            super().close()

    monkeypatch.setattr(writer_module, 'Workbook', TrackingWorkbook)

    config = _basic_config()
    result = _matched_result_with_df(config)
    write_report(result, str(tmp_path / 'report.xlsx'), config)

    assert instances, 'write_report should create a workbook'
    assert all(wb.closed for wb in instances)


def test_write_report_closes_workbook_on_failure(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    """write_report closes its workbook even when saving fails."""
    instances: list = []

    class FailingWorkbook(openpyxl.Workbook):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self.closed = False
            instances.append(self)

        def save(self, *args, **kwargs) -> None:
            raise OSError('simulated save failure')  # noqa: TRY003

        def close(self) -> None:
            self.closed = True
            super().close()

    monkeypatch.setattr(writer_module, 'Workbook', FailingWorkbook)

    config = _basic_config()
    result = _matched_result_with_df(config)
    with pytest.raises(OSError, match='simulated save failure'):
        write_report(result, str(tmp_path / 'report.xlsx'), config)

    assert instances, 'write_report should create a workbook'
    assert all(wb.closed for wb in instances)


# ---------------------------------------------------------------------------
# Unchanged Records — configured columns only (keys + File B comparisons)
# ---------------------------------------------------------------------------
# Generic inventory-style fixture. Column names are arbitrary test data, not
# production constants, so the behavior is verified generically.
_INVENTORY_KEYS = [('SR NO', 'SR')]
_INVENTORY_COMPARE = [
    'on_hand',
    'reserved',
    'available',
    'reorder_point',
    'below_reorder',
    'unit_cost',
    'days_of_cover',
]
_INVENTORY_UNMAPPED = ['snapshot_date', 'warehouse', 'sku']


def _inventory_scenario() -> tuple[ComparisonResult, ComparisonConfig, pd.DataFrame, pd.DataFrame]:
    """One matched record with keys, comparison columns, and unmapped columns."""
    row_a: dict = {'SR NO': '1000'}
    row_b: dict = {'SR': '1000'}
    for i, col in enumerate(_INVENTORY_COMPARE):
        row_a[col] = str(100 + i)
        row_b[col] = str(100 + i)
    for col in _INVENTORY_UNMAPPED:
        row_a[col] = f'a_{col}'
        row_b[col] = f'b_{col}'
    df_a = pd.DataFrame([row_a])
    df_b = pd.DataFrame([row_b])

    config = _config(
        key_pairs=[ColumnPair(fa, fb) for fa, fb in _INVENTORY_KEYS],
        compare_pairs=[ColumnPair(c, c) for c in _INVENTORY_COMPARE],
    )
    result = ComparisonResult(
        matched=[
            RowDiff(
                row_key="SR NO='1000'",
                status=DiffStatus.MATCHED,
                cell_diffs=[],
                index_a=0,
                index_b=0,
                excel_row_a=2,
                excel_row_b=2,
                key_values_a=('1000',),
                key_values_b=('1000',),
            )
        ],
        mismatched=[],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=SchemaDiff(
            common_columns=['SR NO', 'SR', *_INVENTORY_COMPARE],
            only_in_a=list(_INVENTORY_UNMAPPED),
            only_in_b=list(_INVENTORY_UNMAPPED),
            key_columns=['SR NO', 'SR'],
            compare_columns=list(_INVENTORY_COMPARE),
        ),
        summary={'matched_records': 1, 'total_rows_file_a': 1, 'total_rows_file_b': 1},
    )
    return result, config, df_a, df_b


def test_writer_unchanged_records_only_configured_columns(tmp_path: Path):
    """Unchanged Records shows exactly keys + File B comparisons, in order."""
    result, config, df_a, df_b = _inventory_scenario()
    output_path = tmp_path / 'report.xlsx'
    write_report(result, str(output_path), config, df_a=df_a, df_b=df_b)

    workbook = openpyxl.load_workbook(output_path)
    ws = workbook['Unchanged Records']
    headers = [cell.value for cell in ws[1]]

    expected_headers = (
        ['File A – SR NO', 'File B – SR']
        + [f'File B – {c}' for c in _INVENTORY_COMPARE]
        + ['Record Status', 'Excel Row (File A)', 'Excel Row (File B)']
    )
    assert headers == expected_headers

    # Unmapped File A / File B columns are not displayed.
    for col in _INVENTORY_UNMAPPED:
        assert f'File A – {col}' not in headers
        assert f'File B – {col}' not in headers

    data_rows = list(ws.iter_rows(min_row=2, values_only=True))
    assert len(data_rows) == 1
    row = data_rows[0]
    assert row[0] == '1000'  # File A – SR NO
    assert row[1] == '1000'  # File B – SR
    assert list(row[2 : 2 + len(_INVENTORY_COMPARE)]) == [
        str(100 + i) for i in range(len(_INVENTORY_COMPARE))
    ]
    assert row[-3] == 'No changes'
    assert row[-2] == 2
    assert row[-1] == 2
