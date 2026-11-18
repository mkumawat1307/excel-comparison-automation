"""Unit tests for data models."""

from excel_compare.models import (
    CellDiff,
    ComparisonResult,
    DataQualityIssue,
    DiffStatus,
    RowDiff,
    SchemaDiff,
)


def test_diff_status_values() -> None:
    assert DiffStatus.MATCHED == 'matched'
    assert DiffStatus.MISMATCHED == 'mismatched'
    assert DiffStatus.ONLY_IN_A == 'only_in_a'
    assert DiffStatus.ONLY_IN_B == 'only_in_b'
    assert DiffStatus.DUPLICATE == 'duplicate'
    assert DiffStatus.MISSING_KEY == 'missing_key'
    assert DiffStatus.MATCHED_NO_FIELDS_COMPARED == 'matched_no_fields_compared'


def test_diff_status_labels_are_plain_language() -> None:
    assert DiffStatus.MATCHED.label == 'No changes'
    assert DiffStatus.MISMATCHED.label == 'Changed'
    assert DiffStatus.ONLY_IN_A.label == 'Missing from File B'
    assert DiffStatus.ONLY_IN_B.label == 'New in File B'
    assert DiffStatus.DUPLICATE.label == 'Duplicate key'
    assert DiffStatus.MISSING_KEY.label == 'Missing key value'
    assert DiffStatus.MATCHED_NO_FIELDS_COMPARED.label == 'Matched, not compared'


def test_cell_diff_immutable() -> None:
    diff = CellDiff(
        column='Amount',
        row_key='1',
        sheet='Sheet1',
        value_a=10,
        value_b=20,
        are_equal=False,
    )
    assert diff.column == 'Amount'
    assert diff.are_equal is False


def test_row_diff_defaults() -> None:
    row = RowDiff(row_key='1', status=DiffStatus.MATCHED)
    assert row.cell_diffs == []


def test_comparison_result_summary_default() -> None:
    result = ComparisonResult(
        matched=[],
        mismatched=[],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=SchemaDiff(
            common_columns=[], only_in_a=[], only_in_b=[], key_columns=[], compare_columns=[]
        ),
    )
    assert result.summary == {}


def test_data_issue_key_values_default_empty() -> None:
    """DataQualityIssue.key_values_a/key_values_b default to empty tuple."""
    issue = DataQualityIssue(
        issue_type='DUPLICATE_KEY',
        source_file='A',
        original_index=0,
        excel_row_number=2,
        key_display="id='1'",
        all_column_values={'id': '1'},
    )
    assert issue.key_values_a == ()
    assert issue.key_values_b == ()


def test_data_issue_key_values_populated() -> None:
    """DataQualityIssue can carry per-file key values."""
    issue = DataQualityIssue(
        issue_type='DUPLICATE_KEY',
        source_file='A',
        original_index=0,
        excel_row_number=2,
        key_display="id='1'",
        all_column_values={'id': '1'},
        key_values_a=('1',),
        key_values_b=(),
    )
    assert issue.key_values_a == ('1',)
    assert issue.key_values_b == ()
