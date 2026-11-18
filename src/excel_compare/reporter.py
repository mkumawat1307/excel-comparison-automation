"""Report assembly.

Builds the six-sheet ComparisonResult from matcher and comparer output.
Original values are preserved throughout; normalized values are used only
for matching identity.
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple

from .models import (
    CellDiff,
    ColumnPair,
    ComparisonConfig,
    ComparisonResult,
    DataQualityIssue,
    DiffStatus,
    MatchResult,
    RowDiff,
    SchemaDiff,
)


def _key_tuple_display(key_values: Tuple[Any, ...], key_pairs: List[ColumnPair]) -> str:
    """Build human-readable key display string (presentation only)."""
    parts = []
    for pair, v in zip(key_pairs, key_values, strict=False):
        parts.append(f'{pair.file_a_column}={v!r}')
        if pair.file_b_column != pair.file_a_column:
            parts.append(f'{pair.file_b_column}={v!r}')
    return ', '.join(parts)


# Presentation-only heading helpers for per-file source columns.
# Original column names are preserved exactly; no normalization is applied.
FILE_COLUMN_SEPARATOR = ' – '


def file_a_column_header(column_name: str) -> str:
    """Return the display heading for a File A source column.

    Preserves the original column name exactly; only the ``File A`` prefix
    and separator are added.
    """
    return f'File A{FILE_COLUMN_SEPARATOR}{column_name}'


def file_b_column_header(column_name: str) -> str:
    """Return the display heading for a File B source column.

    Preserves the original column name exactly; only the ``File B`` prefix
    and separator are added.
    """
    return f'File B{FILE_COLUMN_SEPARATOR}{column_name}'


# Plain-language labels for summary metrics shown to non-technical users.
# Enum/status values and summary dict keys are unchanged for backward
# compatibility; only the display labels are remapped here.
SUMMARY_LABELS: Dict[str, str] = {
    'total_rows_file_a': 'Records in File A',
    'total_rows_file_b': 'Records in File B',
    'matched_records': 'Records with no changes',
    'mismatched_records': 'Records with changes',
    'matched_no_fields_compared': 'Records matched but not compared',
    'only_in_file_a': 'Records missing from File B',
    'only_in_file_b': 'New records in File B',
    'duplicate_key_records': 'Duplicate identification records',
    'missing_key_records': 'Records with missing identification values',
    'common_columns': 'Common columns',
    'columns_only_in_file_a': 'Columns only in File A',
    'columns_only_in_file_b': 'Columns only in File B',
    'file_a_name': 'File A',
    'file_b_name': 'File B',
}


def summary_label(key: str) -> str:
    """Return the plain-language label for a summary metric key."""
    return SUMMARY_LABELS.get(key, key.replace('_', ' ').title())


def build_report(
    match_result: MatchResult,
    cell_diffs_by_pair: List[List[CellDiff]],
    config: ComparisonConfig,
    excel_row_mapping_a: Optional[Dict[int, int]] = None,
    excel_row_mapping_b: Optional[Dict[int, int]] = None,
) -> ComparisonResult:
    """Assemble the six-sheet ComparisonResult.

    Parameters
    ----------
    match_result : MatchResult
        Output from matcher.match_on_keys().
    cell_diffs_by_pair : list[list[CellDiff]]
        Per-cell comparison results for each matched pair, in the same
        order as match_result.matched_pairs.
    config : ComparisonConfig
        Comparison settings.

    Returns
    -------
    ComparisonResult
        All source rows accounted for exactly once across categories.
    """
    key_pairs = config.key_pairs if config.key_pairs else [ColumnPair(k, k) for k in config.keys]
    compare_pairs = (
        config.compare_pairs
        if config.compare_pairs
        else [ColumnPair(c, c) for c in (config.compare_columns or [])]
    )

    matched: List[RowDiff] = []
    mismatched: List[RowDiff] = []
    matched_no_fields: List[RowDiff] = []

    has_compare_pairs = len(compare_pairs) > 0

    for pair_idx, (row_a, row_b) in enumerate(match_result.matched_pairs):
        cell_diffs = cell_diffs_by_pair[pair_idx] if pair_idx < len(cell_diffs_by_pair) else []
        all_equal = all(d.are_equal for d in cell_diffs)

        key_values_a = tuple(row_a[k.file_a_column] for k in key_pairs)
        key_values_b = tuple(row_b[k.file_b_column] for k in key_pairs)
        row_key = _key_tuple_display(key_values_a, key_pairs)

        def _excel_row_a(row) -> Optional[int]:
            if excel_row_mapping_a and hasattr(row, 'name') and row.name is not None:
                return excel_row_mapping_a.get(int(row.name))
            return None

        def _excel_row_b(row) -> Optional[int]:
            if excel_row_mapping_b and hasattr(row, 'name') and row.name is not None:
                return excel_row_mapping_b.get(int(row.name))
            return None

        if not has_compare_pairs:
            # No fields compared: do not classify as identical.
            status = DiffStatus.MATCHED_NO_FIELDS_COMPARED
            matched_no_fields.append(
                RowDiff(
                    row_key=row_key,
                    status=status,
                    cell_diffs=list(cell_diffs),
                    index_a=(
                        int(row_a.name)
                        if hasattr(row_a, 'name') and row_a.name is not None
                        else None
                    ),
                    index_b=(
                        int(row_b.name)
                        if hasattr(row_b, 'name') and row_b.name is not None
                        else None
                    ),
                    excel_row_a=_excel_row_a(row_a),
                    excel_row_b=_excel_row_b(row_b),
                    key_values_a=key_values_a,
                    key_values_b=key_values_b,
                )
            )
        elif all_equal:
            matched.append(
                RowDiff(
                    row_key=row_key,
                    status=DiffStatus.MATCHED,
                    cell_diffs=list(cell_diffs),
                    index_a=(
                        int(row_a.name)
                        if hasattr(row_a, 'name') and row_a.name is not None
                        else None
                    ),
                    index_b=(
                        int(row_b.name)
                        if hasattr(row_b, 'name') and row_b.name is not None
                        else None
                    ),
                    excel_row_a=_excel_row_a(row_a),
                    excel_row_b=_excel_row_b(row_b),
                    key_values_a=key_values_a,
                    key_values_b=key_values_b,
                )
            )
        else:
            mismatched.append(
                RowDiff(
                    row_key=row_key,
                    status=DiffStatus.MISMATCHED,
                    cell_diffs=list(cell_diffs),
                    index_a=(
                        int(row_a.name)
                        if hasattr(row_a, 'name') and row_a.name is not None
                        else None
                    ),
                    index_b=(
                        int(row_b.name)
                        if hasattr(row_b, 'name') and row_b.name is not None
                        else None
                    ),
                    excel_row_a=_excel_row_a(row_a),
                    excel_row_b=_excel_row_b(row_b),
                    key_values_a=key_values_a,
                    key_values_b=key_values_b,
                )
            )

    only_in_a: List[RowDiff] = []
    for row in match_result.only_in_a_unique:
        key_values_a = tuple(row[k.file_a_column] for k in key_pairs)
        row_key = _key_tuple_display(key_values_a, key_pairs)
        only_in_a.append(
            RowDiff(
                row_key=row_key,
                status=DiffStatus.ONLY_IN_A,
                index_a=(int(row.name) if hasattr(row, 'name') and row.name is not None else None),
                excel_row_a=(
                    excel_row_mapping_a.get(int(row.name))
                    if excel_row_mapping_a and hasattr(row, 'name') and row.name is not None
                    else None
                ),
                key_values_a=key_values_a,
                key_values_b=(),
            )
        )

    only_in_b: List[RowDiff] = []
    for row in match_result.only_in_b_unique:
        key_values_b = tuple(row[k.file_b_column] for k in key_pairs)
        row_key = _key_tuple_display(key_values_b, key_pairs)
        only_in_b.append(
            RowDiff(
                row_key=row_key,
                status=DiffStatus.ONLY_IN_B,
                index_b=(int(row.name) if hasattr(row, 'name') and row.name is not None else None),
                excel_row_b=(
                    excel_row_mapping_b.get(int(row.name))
                    if excel_row_mapping_b and hasattr(row, 'name') and row.name is not None
                    else None
                ),
                key_values_a=(),
                key_values_b=key_values_b,
            )
        )

    data_quality_issues: List[DataQualityIssue] = []

    for rec in match_result.duplicate_source_records_a:
        data_quality_issues.append(
            DataQualityIssue(
                issue_type='DUPLICATE_KEY',
                source_file=rec.source_file,
                original_index=rec.original_index,
                excel_row_number=rec.excel_row_number,
                key_display=rec.key_display,
                all_column_values=rec.all_column_values,
                key_values_a=rec.key_values,
                key_values_b=(),
            )
        )

    for rec in match_result.duplicate_source_records_b:
        data_quality_issues.append(
            DataQualityIssue(
                issue_type='DUPLICATE_KEY',
                source_file=rec.source_file,
                original_index=rec.original_index,
                excel_row_number=rec.excel_row_number,
                key_display=rec.key_display,
                all_column_values=rec.all_column_values,
                key_values_a=(),
                key_values_b=rec.key_values,
            )
        )

    for rec in match_result.missing_key_records_a:
        data_quality_issues.append(
            DataQualityIssue(
                issue_type='MISSING_KEY',
                source_file=rec.source_file,
                original_index=rec.original_index,
                excel_row_number=rec.excel_row_number,
                key_display=rec.key_display,
                all_column_values=rec.all_column_values,
                key_values_a=rec.key_values,
                key_values_b=(),
            )
        )

    for rec in match_result.missing_key_records_b:
        data_quality_issues.append(
            DataQualityIssue(
                issue_type='MISSING_KEY',
                source_file=rec.source_file,
                original_index=rec.original_index,
                excel_row_number=rec.excel_row_number,
                key_display=rec.key_display,
                all_column_values=rec.all_column_values,
                key_values_a=(),
                key_values_b=rec.key_values,
            )
        )

    total_a = (
        len(match_result.matched_pairs)
        + len(match_result.only_in_a_unique)
        + len(match_result.duplicate_source_records_a)
        + len(match_result.missing_key_records_a)
    )
    total_b = (
        len(match_result.matched_pairs)
        + len(match_result.only_in_b_unique)
        + len(match_result.duplicate_source_records_b)
        + len(match_result.missing_key_records_b)
    )

    duplicate_count = sum(1 for i in data_quality_issues if i.issue_type == 'DUPLICATE_KEY')
    missing_count = sum(1 for i in data_quality_issues if i.issue_type == 'MISSING_KEY')

    schema_diff = match_result.schema_diff
    if schema_diff is None:
        schema_diff = SchemaDiff(
            common_columns=[],
            only_in_a=[],
            only_in_b=[],
            key_columns=[],
            compare_columns=[],
            exact_name_key_pairs=[],
            exact_name_compare_pairs=[],
            renamed_key_pairs=[],
            renamed_compare_pairs=[],
            unmapped_file_a_columns=[],
            unmapped_file_b_columns=[],
        )

    summary: Dict[str, Any] = {
        'total_rows_file_a': total_a,
        'total_rows_file_b': total_b,
        'matched_records': len(matched),
        'mismatched_records': len(mismatched),
        'matched_no_fields_compared': len(matched_no_fields),
        'only_in_file_a': len(only_in_a),
        'only_in_file_b': len(only_in_b),
        'duplicate_key_records': duplicate_count,
        'missing_key_records': missing_count,
        'duplicate_groups_file_a': len(match_result.duplicate_groups_a),
        'duplicate_groups_file_b': len(match_result.duplicate_groups_b),
        'common_columns': len(schema_diff.common_columns),
        'columns_only_in_file_a': len(schema_diff.only_in_a),
        'columns_only_in_file_b': len(schema_diff.only_in_b),
    }

    return ComparisonResult(
        matched=matched,
        mismatched=mismatched,
        only_in_a=only_in_a,
        only_in_b=only_in_b,
        data_quality_issues=data_quality_issues,
        schema_diff=schema_diff,
        summary=summary,
    )
