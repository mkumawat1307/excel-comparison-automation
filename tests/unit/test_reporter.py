"""Unit tests for the reporter module (Phase 4)."""

from __future__ import annotations

import pandas as pd

from excel_compare.models import (
    CellDiff,
    ColumnPair,
    ComparisonConfig,
    DataQualityIssue,
    DuplicateSourceRecord,
    MatchResult,
    MissingKeyRecord,
    SchemaDiff,
)
from excel_compare.reporter import build_report, summary_label


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


def test_report_assembly_basic():
    row_a = _row({'id': '1', 'amount': '100', 'name': 'Alice'})
    row_b = _row({'id': '1', 'amount': '100', 'name': 'Alice'})
    config = _config(keys=['id'], compare_columns=['amount', 'name'])

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
            common_columns=['id', 'amount', 'name'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=['amount', 'name'],
        ),
    )

    cell_diffs = [
        [
            CellDiff(
                column='amount',
                row_key='id=1',
                sheet='',
                value_a='100',
                value_b='100',
                are_equal=True,
            ),
            CellDiff(
                column='name',
                row_key='id=1',
                sheet='',
                value_a='Alice',
                value_b='Alice',
                are_equal=True,
            ),
        ]
    ]

    result = build_report(match_result, cell_diffs, config)

    assert len(result.matched) == 1
    assert len(result.mismatched) == 0
    assert len(result.only_in_a) == 0
    assert len(result.only_in_b) == 0
    assert len(result.data_quality_issues) == 0
    assert result.summary['matched_records'] == 1


def test_report_mismatched_record():
    row_a = _row({'id': '1', 'amount': '100', 'name': 'Alice'})
    row_b = _row({'id': '1', 'amount': '200', 'name': 'Alice'})
    config = _config(keys=['id'], compare_columns=['amount', 'name'])

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
            common_columns=['id', 'amount', 'name'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=['amount', 'name'],
        ),
    )

    cell_diffs = [
        [
            CellDiff(
                column='amount',
                row_key='id=1',
                sheet='',
                value_a='100',
                value_b='200',
                are_equal=False,
            ),
            CellDiff(
                column='name',
                row_key='id=1',
                sheet='',
                value_a='Alice',
                value_b='Alice',
                are_equal=True,
            ),
        ]
    ]

    result = build_report(match_result, cell_diffs, config)

    assert len(result.matched) == 0
    assert len(result.mismatched) == 1
    assert result.mismatched[0].cell_diffs[0].are_equal is False
    assert result.mismatched[0].cell_diffs[0].value_a == '100'
    assert result.mismatched[0].cell_diffs[0].value_b == '200'


def test_report_duplicate_records_not_classified_as_missing():
    """Duplicate records have issue_type DUPLICATE_KEY, not MISSING_KEY."""
    rec = DuplicateSourceRecord(
        source_file='A',
        original_index=5,
        excel_row_number=7,
        key_values=('1',),
        key_display="id='1'",
        group_occurrence_count=2,
        all_column_values={'id': '1', 'name': 'Alice'},
    )
    issue = DataQualityIssue(
        issue_type='DUPLICATE_KEY',
        source_file=rec.source_file,
        original_index=rec.original_index,
        excel_row_number=rec.excel_row_number,
        key_display=rec.key_display,
        all_column_values=rec.all_column_values,
    )
    assert issue.issue_type == 'DUPLICATE_KEY'
    assert issue.issue_type != 'MISSING_KEY'


def test_report_missing_key_records():
    """Missing key records have issue_type MISSING_KEY."""
    rec = MissingKeyRecord(
        source_file='B',
        original_index=3,
        excel_row_number=5,
        key_values=('',),
        key_display="id=''",
        all_column_values={'id': '', 'name': 'Bob'},
    )
    issue = DataQualityIssue(
        issue_type='MISSING_KEY',
        source_file=rec.source_file,
        original_index=rec.original_index,
        excel_row_number=rec.excel_row_number,
        key_display=rec.key_display,
        all_column_values=rec.all_column_values,
    )
    assert issue.issue_type == 'MISSING_KEY'
    assert issue.issue_type != 'DUPLICATE_KEY'


def test_report_schema_diff_preserved():
    row_a = _row({'id': '1', 'amount': '100'})
    row_b = _row({'id': '1', 'amount': '100'})
    config = _config(keys=['id'], compare_columns=['amount'])
    schema = SchemaDiff(
        common_columns=['id', 'amount'],
        only_in_a=['extra_a'],
        only_in_b=['extra_b'],
        key_columns=['id'],
        compare_columns=['amount'],
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
        schema_diff=schema,
    )

    result = build_report(match_result, [], config)
    assert result.schema_diff.common_columns == ['id', 'amount']
    assert result.schema_diff.only_in_a == ['extra_a']
    assert result.schema_diff.only_in_b == ['extra_b']
    assert result.schema_diff.compare_columns == ['amount']


def test_report_only_in_a_b():
    row_a = _row({'id': '1', 'name': 'Alice'})
    row_b = _row({'id': '2', 'name': 'Bob'})
    config = _config(keys=['id'], compare_columns=['name'])

    match_result = MatchResult(
        matched_pairs=[],
        only_in_a_unique=[row_a],
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

    result = build_report(match_result, [], config)
    assert len(result.only_in_a) == 1
    assert result.only_in_a[0].row_key == "id='1'"
    assert len(result.only_in_b) == 1
    assert result.only_in_b[0].row_key == "id='2'"


def test_row_accounting_in_summary():
    """Summary counts reflect all rows accounted for."""
    row_a1 = _row({'id': '1', 'amount': '100'})
    row_a2 = _row({'id': '3', 'amount': '300'})
    row_b1 = _row({'id': '1', 'amount': '100'})
    row_b2 = _row({'id': '4', 'amount': '400'})

    dup_rec = DuplicateSourceRecord(
        source_file='A',
        original_index=5,
        excel_row_number=7,
        key_values=('99',),
        key_display="id='99'",
        group_occurrence_count=2,
        all_column_values={'id': '99', 'name': 'Dup'},
    )

    match_result = MatchResult(
        matched_pairs=[(row_a1, row_b1)],
        only_in_a_unique=[row_a2],
        only_in_b_unique=[row_b2],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[dup_rec],
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

    cell_diffs = [
        [
            CellDiff(
                column='amount',
                row_key='id=1',
                sheet='',
                value_a='100',
                value_b='100',
                are_equal=True,
            ),
        ]
    ]

    config = _config(keys=['id'], compare_columns=['amount'])
    result = build_report(match_result, cell_diffs, config)

    assert result.summary['total_rows_file_a'] == 3  # 1 matched + 1 only_in_a + 1 duplicate
    assert result.summary['total_rows_file_b'] == 2  # 1 matched + 1 only_in_b
    assert result.summary['matched_records'] == 1
    assert result.summary['only_in_file_a'] == 1
    assert result.summary['only_in_file_b'] == 1
    assert result.summary['duplicate_key_records'] == 1


def test_report_renamed_key_pairs():
    """Renamed key pairs appear in schema diff and row keys."""
    row_a = _row({'Site ID': 'S1', 'name': 'Alice'})
    row_b = _row({'Site Code': 'S1', 'name': 'Alice'})
    key_pairs = [ColumnPair('Site ID', 'Site Code')]
    compare_pairs = [ColumnPair('name', 'name')]
    config = _config(keys=[], key_pairs=key_pairs, compare_pairs=compare_pairs)
    schema = SchemaDiff(
        common_columns=['name'],
        only_in_a=['Site ID'],
        only_in_b=['Site Code'],
        key_columns=['Site ID'],
        compare_columns=['name'],
        exact_name_key_pairs=[],
        exact_name_compare_pairs=[ColumnPair('name', 'name')],
        renamed_key_pairs=[ColumnPair('Site ID', 'Site Code')],
        renamed_compare_pairs=[],
        unmapped_file_a_columns=[],
        unmapped_file_b_columns=[],
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
        schema_diff=schema,
    )
    cell_diffs = [
        [
            CellDiff(
                column_a='name',
                column_b='name',
                row_key="Site ID='S1', Site Code='S1'",
                sheet='',
                value_a='Alice',
                value_b='Alice',
                are_equal=True,
            )
        ]
    ]
    result = build_report(match_result, cell_diffs, config)
    assert len(result.matched) == 1
    assert result.matched[0].row_key == "Site ID='S1', Site Code='S1'"
    assert result.schema_diff.renamed_key_pairs == [ColumnPair('Site ID', 'Site Code')]


def test_report_renamed_compare_pairs():
    """Renamed compare pairs appear in CellDiff column_a/column_b."""
    row_a = _row({'Make': 'Toyota', 'id': '1'})
    row_b = _row({'Camera Make': 'Toyota', 'id': '1'})
    key_pairs = [ColumnPair('id', 'id')]
    compare_pairs = [ColumnPair('Make', 'Camera Make')]
    config = _config(keys=[], key_pairs=key_pairs, compare_pairs=compare_pairs)
    schema = SchemaDiff(
        common_columns=['id'],
        only_in_a=['Make'],
        only_in_b=['Camera Make'],
        key_columns=['id'],
        compare_columns=[],
        exact_name_key_pairs=[ColumnPair('id', 'id')],
        exact_name_compare_pairs=[],
        renamed_key_pairs=[],
        renamed_compare_pairs=[ColumnPair('Make', 'Camera Make')],
        unmapped_file_a_columns=['Make'],
        unmapped_file_b_columns=['Camera Make'],
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
        schema_diff=schema,
    )
    cell_diffs = [
        [
            CellDiff(
                column_a='Make',
                column_b='Camera Make',
                row_key="id='1'",
                sheet='',
                value_a='Toyota',
                value_b='Toyota',
                are_equal=True,
            )
        ]
    ]
    result = build_report(match_result, cell_diffs, config)
    assert len(result.matched) == 1
    assert result.matched[0].cell_diffs[0].column_a == 'Make'
    assert result.matched[0].cell_diffs[0].column_b == 'Camera Make'


def test_report_empty_compare_pairs_matched_no_fields():
    """Empty compare_pairs → MATCHED_NO_FIELDS_COMPARED, not MATCHED."""
    row_a = _row({'id': '1', 'name': 'Alice'})
    row_b = _row({'id': '1', 'name': 'Bob'})
    key_pairs = [ColumnPair('id', 'id')]
    config = _config(keys=[], key_pairs=key_pairs, compare_pairs=[])
    schema = SchemaDiff(
        common_columns=['id', 'name'],
        only_in_a=[],
        only_in_b=[],
        key_columns=['id'],
        compare_columns=[],
        exact_name_key_pairs=[ColumnPair('id', 'id')],
        exact_name_compare_pairs=[],
        renamed_key_pairs=[],
        renamed_compare_pairs=[],
        unmapped_file_a_columns=['name'],
        unmapped_file_b_columns=['name'],
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
        schema_diff=schema,
    )
    result = build_report(match_result, [], config)
    assert len(result.matched) == 0
    assert len(result.mismatched) == 0
    # matched_no_fields_compared is tracked separately
    assert result.summary['matched_no_fields_compared'] == 1


def test_report_unmapped_columns_not_in_mismatch_count():
    """Unmapped columns do not appear in mismatch counts."""
    row_a = _row({'id': '1', 'extra_a': 'x'})
    row_b = _row({'id': '1', 'extra_b': 'y'})
    key_pairs = [ColumnPair('id', 'id')]
    config = _config(keys=[], key_pairs=key_pairs, compare_pairs=[])
    schema = SchemaDiff(
        common_columns=['id'],
        only_in_a=['extra_a'],
        only_in_b=['extra_b'],
        key_columns=['id'],
        compare_columns=[],
        exact_name_key_pairs=[ColumnPair('id', 'id')],
        exact_name_compare_pairs=[],
        renamed_key_pairs=[],
        renamed_compare_pairs=[],
        unmapped_file_a_columns=['extra_a'],
        unmapped_file_b_columns=['extra_b'],
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
        schema_diff=schema,
    )
    result = build_report(match_result, [], config)
    assert result.summary['mismatched_records'] == 0
    assert result.summary['matched_no_fields_compared'] == 1
    assert 'extra_a' in result.schema_diff.unmapped_file_a_columns
    assert 'extra_b' in result.schema_diff.unmapped_file_b_columns


def test_build_report_preserves_key_values_single_key():
    """build_report preserves per-column key values in RowDiff.key_values_a/b."""
    row_a = _row({'id': '101', 'name': 'Alice'})
    row_b = _row({'id': '101', 'name': 'Alice'})
    key_pairs = [ColumnPair('id', 'id')]
    compare_pairs = [ColumnPair('name', 'name')]
    config = _config(keys=[], key_pairs=key_pairs, compare_pairs=compare_pairs)

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

    cell_diffs = [
        [
            CellDiff(
                column_a='name',
                column_b='name',
                row_key="id='101'",
                sheet='',
                value_a='Alice',
                value_b='Alice',
                are_equal=True,
            )
        ]
    ]

    result = build_report(match_result, cell_diffs, config)

    assert len(result.matched) == 1
    assert result.matched[0].key_values_a == ('101',)
    assert result.matched[0].key_values_b == ('101',)
    # row_key output remains unchanged
    assert result.matched[0].row_key == "id='101'"


def test_build_report_preserves_key_values_composite_key():
    """Composite keys preserve every individual key value."""
    row_a = _row({'dept': 'HR', 'emp_id': 'E001'})
    row_b = _row({'dept': 'HR', 'emp_id': 'E001'})
    key_pairs = [ColumnPair('dept', 'dept'), ColumnPair('emp_id', 'emp_id')]
    compare_pairs = [ColumnPair('dept', 'dept')]
    config = _config(keys=[], key_pairs=key_pairs, compare_pairs=compare_pairs)

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
            common_columns=['dept', 'emp_id'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['dept', 'emp_id'],
            compare_columns=[],
        ),
    )

    result = build_report(match_result, [], config)

    assert len(result.matched) == 1
    assert result.matched[0].key_values_a == ('HR', 'E001')
    assert result.matched[0].key_values_b == ('HR', 'E001')
    assert result.matched[0].row_key == "dept='HR', emp_id='E001'"


def test_build_report_preserves_key_values_renamed_columns():
    """Differently named key columns produce correct per-file key values."""
    row_a = _row({'Site ID': 'S1', 'name': 'Alice'})
    row_b = _row({'Site Code': 'S1', 'name': 'Alice'})
    key_pairs = [ColumnPair('Site ID', 'Site Code')]
    compare_pairs = [ColumnPair('name', 'name')]
    config = _config(keys=[], key_pairs=key_pairs, compare_pairs=compare_pairs)

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
            common_columns=['name'],
            only_in_a=['Site ID'],
            only_in_b=['Site Code'],
            key_columns=['Site ID'],
            compare_columns=['name'],
            exact_name_key_pairs=[],
            exact_name_compare_pairs=compare_pairs,
            renamed_key_pairs=key_pairs,
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
                row_key="Site ID='S1', Site Code='S1'",
                sheet='',
                value_a='Alice',
                value_b='Alice',
                are_equal=True,
            )
        ]
    ]

    result = build_report(match_result, cell_diffs, config)

    assert len(result.matched) == 1
    assert result.matched[0].key_values_a == ('S1',)
    assert result.matched[0].key_values_b == ('S1',)
    assert result.matched[0].row_key == "Site ID='S1', Site Code='S1'"


def test_build_report_summary_matched_excludes_not_compared():
    """Records matched but not compared are never counted as unchanged."""
    row_a = _row({'id': '1', 'name': 'Alice'})
    row_b = _row({'id': '1', 'name': 'Alice'})

    # No compare columns -> every matched pair is "matched, not compared"
    # and must not inflate the unchanged count.
    key_pairs = [ColumnPair('id', 'id')]
    config = _config(keys=[], key_pairs=key_pairs, compare_pairs=[])

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
            compare_columns=[],
        ),
    )

    result = build_report(match_result, [], config)

    assert len(result.matched) == 0
    assert result.summary['matched_records'] == 0
    assert result.summary['matched_no_fields_compared'] == 1
    assert result.summary['mismatched_records'] == 0


def test_build_report_summary_only_in_and_quality_counts():
    """Summary counts only-in records, duplicates and missing keys."""
    row_a = _row({'id': '1'})
    row_b = _row({'id': '2'})
    key_pairs = [ColumnPair('id', 'id')]
    config = _config(keys=[], key_pairs=key_pairs, compare_pairs=[])

    dup_rec = DuplicateSourceRecord(
        source_file='File A',
        original_index=0,
        excel_row_number=2,
        key_values=('1',),
        key_display="id='1'",
        group_occurrence_count=2,
        all_column_values={'id': '1'},
    )
    missing_rec = MissingKeyRecord(
        source_file='File B',
        original_index=3,
        excel_row_number=5,
        key_values=('',),
        key_display="id=''",
        all_column_values={'id': ''},
    )

    match_result = MatchResult(
        matched_pairs=[],
        only_in_a_unique=[row_a],
        only_in_b_unique=[row_b],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[dup_rec],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[missing_rec],
        schema_diff=SchemaDiff(
            common_columns=['id'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=[],
        ),
    )

    result = build_report(match_result, [], config)

    assert result.summary['only_in_file_a'] == 1
    assert result.summary['only_in_file_b'] == 1
    assert result.summary['duplicate_key_records'] == 1
    assert result.summary['missing_key_records'] == 1
    assert result.summary['total_rows_file_a'] == 2  # only_in_a + duplicate
    assert result.summary['total_rows_file_b'] == 2  # only_in_b + missing key


def test_summary_label_is_plain_language():
    """summary_label returns client-friendly labels for summary keys."""
    assert summary_label('matched_records') == 'Records with no changes'
    assert summary_label('mismatched_records') == 'Records with changes'
    assert summary_label('matched_no_fields_compared') == 'Records matched but not compared'
    assert summary_label('duplicate_key_records') == 'Duplicate identification records'
    assert summary_label('missing_key_records') == 'Records with missing identification values'
    assert summary_label('only_in_file_a') == 'Records missing from File B'
    assert summary_label('only_in_file_b') == 'New records in File B'
    # Unknown keys fall back to title case
    assert summary_label('weird_metric') == 'Weird Metric'


# ---------------------------------------------------------------------------
# Per-file key value regression tests for build_report
# ---------------------------------------------------------------------------


def test_build_report_only_in_a_populates_file_a_keys_only():
    """only_in_a RowDiff has key_values_a populated, key_values_b empty."""
    row_a = _row({'id': '9', 'name': 'Orphan'})
    key_pairs = [ColumnPair('id', 'id')]
    config = _config(keys=[], key_pairs=key_pairs, compare_pairs=[])

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
            compare_columns=[],
        ),
    )

    result = build_report(match_result, [], config)

    assert len(result.only_in_a) == 1
    assert result.only_in_a[0].key_values_a == ('9',)
    assert result.only_in_a[0].key_values_b == ()


def test_build_report_only_in_b_populates_file_b_keys_only():
    """only_in_b RowDiff has key_values_b populated, key_values_a empty."""
    row_b = _row({'id': '9', 'name': 'Newbie'})
    key_pairs = [ColumnPair('id', 'id')]
    config = _config(keys=[], key_pairs=key_pairs, compare_pairs=[])

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
            compare_columns=[],
        ),
    )

    result = build_report(match_result, [], config)

    assert len(result.only_in_b) == 1
    assert result.only_in_b[0].key_values_a == ()
    assert result.only_in_b[0].key_values_b == ('9',)


def test_build_report_data_quality_preserves_per_file_keys():
    """DataQualityIssue preserves per-file key values from source records."""
    dup_rec_a = DuplicateSourceRecord(
        source_file='A',
        original_index=0,
        excel_row_number=2,
        key_values=('1',),
        key_display="id='1'",
        group_occurrence_count=2,
        all_column_values={'id': '1'},
    )
    missing_rec_b = MissingKeyRecord(
        source_file='B',
        original_index=3,
        excel_row_number=5,
        key_values=('',),
        key_display="id=''",
        all_column_values={'id': ''},
    )

    match_result = MatchResult(
        matched_pairs=[],
        only_in_a_unique=[],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[dup_rec_a],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[missing_rec_b],
        schema_diff=SchemaDiff(
            common_columns=['id'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=[],
        ),
    )

    result = build_report(
        match_result, [], config=_config(key_pairs=[ColumnPair('id', 'id')], compare_pairs=[])
    )

    # Find the DUPLICATE_KEY issue (from file A)
    dup_issues = [i for i in result.data_quality_issues if i.issue_type == 'DUPLICATE_KEY']
    assert len(dup_issues) == 1
    assert dup_issues[0].key_values_a == ('1',)
    assert dup_issues[0].key_values_b == ()

    # Find the MISSING_KEY issue (from file B)
    missing_issues = [i for i in result.data_quality_issues if i.issue_type == 'MISSING_KEY']
    assert len(missing_issues) == 1
    assert missing_issues[0].key_values_a == ()
    assert missing_issues[0].key_values_b == ('',)


# ---------------------------------------------------------------------------
# Record ordering regression tests (Issue 3)
# Verify build_report preserves matcher output order.
# ---------------------------------------------------------------------------


def test_report_matched_records_preserve_order():
    """Matched records order follows matcher output order."""
    row_a1 = _row({'id': 'C', 'name': 'Carol'})
    row_b1 = _row({'id': 'C', 'name': 'Carol'})
    row_a2 = _row({'id': 'A', 'name': 'Alice'})
    row_b2 = _row({'id': 'A', 'name': 'Alice'})
    row_a3 = _row({'id': 'B', 'name': 'Bob'})
    row_b3 = _row({'id': 'B', 'name': 'Bob'})
    key_pairs = [ColumnPair('id', 'id')]
    compare_pairs = [ColumnPair('name', 'name')]
    config = _config(keys=[], key_pairs=key_pairs, compare_pairs=compare_pairs)

    match_result = MatchResult(
        matched_pairs=[(row_a1, row_b1), (row_a2, row_b2), (row_a3, row_b3)],
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

    cell_diffs = [
        [
            CellDiff(
                column_a='name',
                column_b='name',
                row_key="id='C'",
                sheet='',
                value_a='Carol',
                value_b='Carol',
                are_equal=True,
            )
        ],
        [
            CellDiff(
                column_a='name',
                column_b='name',
                row_key="id='A'",
                sheet='',
                value_a='Alice',
                value_b='Alice',
                are_equal=True,
            )
        ],
        [
            CellDiff(
                column_a='name',
                column_b='name',
                row_key="id='B'",
                sheet='',
                value_a='Bob',
                value_b='Bob',
                are_equal=True,
            )
        ],
    ]

    result = build_report(match_result, cell_diffs, config)

    assert len(result.matched) == 3
    assert result.matched[0].key_values_a == ('C',)
    assert result.matched[1].key_values_a == ('A',)
    assert result.matched[2].key_values_a == ('B',)


def test_report_only_in_a_preserve_order():
    """Only-in-A records order follows matcher output order."""
    row_a1 = _row({'id': 'Z', 'name': 'Zara'})
    row_a2 = _row({'id': 'A', 'name': 'Alice'})
    row_a3 = _row({'id': 'M', 'name': 'Mallory'})
    key_pairs = [ColumnPair('id', 'id')]
    config = _config(keys=[], key_pairs=key_pairs, compare_pairs=[])

    match_result = MatchResult(
        matched_pairs=[],
        only_in_a_unique=[row_a1, row_a2, row_a3],
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
            compare_columns=[],
        ),
    )

    result = build_report(match_result, [], config)

    assert len(result.only_in_a) == 3
    assert result.only_in_a[0].key_values_a == ('Z',)
    assert result.only_in_a[1].key_values_a == ('A',)
    assert result.only_in_a[2].key_values_a == ('M',)


def test_report_only_in_b_preserve_order():
    """Only-in-B records order follows matcher output order."""
    row_b1 = _row({'id': 'Z', 'name': 'Zara'})
    row_b2 = _row({'id': 'A', 'name': 'Alice'})
    row_b3 = _row({'id': 'M', 'name': 'Mallory'})
    key_pairs = [ColumnPair('id', 'id')]
    config = _config(keys=[], key_pairs=key_pairs, compare_pairs=[])

    match_result = MatchResult(
        matched_pairs=[],
        only_in_a_unique=[],
        only_in_b_unique=[row_b1, row_b2, row_b3],
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
            compare_columns=[],
        ),
    )

    result = build_report(match_result, [], config)

    assert len(result.only_in_b) == 3
    assert result.only_in_b[0].key_values_b == ('Z',)
    assert result.only_in_b[1].key_values_b == ('A',)
    assert result.only_in_b[2].key_values_b == ('M',)


def test_report_data_quality_preserve_order():
    """Data quality issues preserve matcher output order (duplicates then missing keys)."""
    dup_rec_a1 = DuplicateSourceRecord(
        source_file='A',
        original_index=0,
        excel_row_number=2,
        key_values=('DUP1',),
        key_display="id='DUP1'",
        group_occurrence_count=2,
        all_column_values={'id': 'DUP1'},
    )
    dup_rec_a2 = DuplicateSourceRecord(
        source_file='A',
        original_index=2,
        excel_row_number=4,
        key_values=('DUP2',),
        key_display="id='DUP2'",
        group_occurrence_count=1,
        all_column_values={'id': 'DUP2'},
    )
    missing_rec_b1 = MissingKeyRecord(
        source_file='B',
        original_index=5,
        excel_row_number=7,
        key_values=('',),
        key_display="id=''",
        all_column_values={'id': ''},
    )
    missing_rec_b2 = MissingKeyRecord(
        source_file='B',
        original_index=6,
        excel_row_number=8,
        key_values=('',),
        key_display="id=''",
        all_column_values={'id': ''},
    )

    match_result = MatchResult(
        matched_pairs=[],
        only_in_a_unique=[],
        only_in_b_unique=[],
        duplicate_groups_a=[],
        duplicate_groups_b=[],
        duplicate_source_records_a=[dup_rec_a1, dup_rec_a2],
        duplicate_source_records_b=[],
        missing_key_records_a=[],
        missing_key_records_b=[missing_rec_b1, missing_rec_b2],
        schema_diff=SchemaDiff(
            common_columns=['id'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['id'],
            compare_columns=[],
        ),
    )

    result = build_report(
        match_result, [], config=_config(key_pairs=[ColumnPair('id', 'id')], compare_pairs=[])
    )

    assert len(result.data_quality_issues) == 4
    # Duplicates (File A) appear first, in matcher order
    assert result.data_quality_issues[0].issue_type == 'DUPLICATE_KEY'
    assert result.data_quality_issues[0].key_values_a == ('DUP1',)
    assert result.data_quality_issues[1].issue_type == 'DUPLICATE_KEY'
    assert result.data_quality_issues[1].key_values_a == ('DUP2',)
    # Missing keys (File B) appear next, in matcher order
    assert result.data_quality_issues[2].issue_type == 'MISSING_KEY'
    assert result.data_quality_issues[2].key_values_b == ('',)
    assert result.data_quality_issues[3].issue_type == 'MISSING_KEY'
    assert result.data_quality_issues[3].key_values_b == ('',)
