"""Typed data models for the Excel comparison application."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd


class DiffStatus(str, Enum):
    MATCHED = 'matched'
    MISMATCHED = 'mismatched'
    ONLY_IN_A = 'only_in_a'
    ONLY_IN_B = 'only_in_b'
    DUPLICATE = 'duplicate'
    MISSING_KEY = 'missing_key'
    MATCHED_NO_FIELDS_COMPARED = 'matched_no_fields_compared'

    @property
    def label(self) -> str:
        """Plain-language label for display to non-technical users.

        The enum value (e.g. 'matched') is preserved for data interchange;
        only the display label changes here.
        """
        return {
            DiffStatus.MATCHED: 'No changes',
            DiffStatus.MISMATCHED: 'Changed',
            DiffStatus.ONLY_IN_A: 'Missing from File B',
            DiffStatus.ONLY_IN_B: 'New in File B',
            DiffStatus.DUPLICATE: 'Duplicate key',
            DiffStatus.MISSING_KEY: 'Missing key value',
            DiffStatus.MATCHED_NO_FIELDS_COMPARED: 'Matched, not compared',
        }[self]


@dataclass(frozen=True)
class ColumnPair:
    file_a_column: str
    file_b_column: str

    def __post_init__(self) -> None:
        if not self.file_a_column:
            raise ValueError('file_a_column must not be empty')  # noqa: TRY003
        if not self.file_b_column:
            raise ValueError('file_b_column must not be empty')  # noqa: TRY003


@dataclass(frozen=True)
class CellDiff:
    column_a: Optional[str] = None
    column_b: Optional[str] = None
    row_key: str = ''
    sheet: str = ''
    value_a: Any = None
    value_b: Any = None
    are_equal: bool = False
    index_a: Optional[int] = None
    index_b: Optional[int] = None
    excel_row_a: Optional[int] = None
    excel_row_b: Optional[int] = None
    column: Optional[str] = None

    def __post_init__(self) -> None:
        # Backward-compatible: if column is set and column_a is empty, use it.
        if self.column is not None and self.column_a is None:
            object.__setattr__(self, 'column_a', self.column)
        if self.column is not None and self.column_b is None:
            object.__setattr__(self, 'column_b', self.column)


@dataclass(frozen=True)
class RowDiff:
    row_key: str
    status: DiffStatus
    cell_diffs: List[CellDiff] = field(default_factory=list)
    index_a: Optional[int] = None
    index_b: Optional[int] = None
    excel_row_a: Optional[int] = None
    excel_row_b: Optional[int] = None
    key_values_a: tuple = ()
    key_values_b: tuple = ()


@dataclass(frozen=True)
class DuplicateRecord:
    row_key: str
    sheet: str
    occurrences: int
    row_indices: List[int] = field(default_factory=list)


@dataclass(frozen=True)
class ComparisonConfig:
    # Legacy (deprecated, same-name only)
    keys: List[str] = field(default_factory=list)
    compare_columns: Optional[List[str]] = None
    numeric_columns: Optional[List[str]] = None
    date_columns: Optional[List[str]] = None

    # New pair-based (primary)
    key_pairs: List[ColumnPair] = field(default_factory=list)
    compare_pairs: List[ColumnPair] = field(default_factory=list)
    numeric_pairs: List[ColumnPair] = field(default_factory=list)
    date_pairs: List[ColumnPair] = field(default_factory=list)

    case_sensitive: bool = False
    normalize_whitespace: bool = True
    numeric_tolerance: float = 1e-9
    duplicate_policy: str = 'error'
    max_file_size_mb: int = 50


@dataclass(frozen=True)
class LoadedSheet:
    """Metadata-aware container for a loaded worksheet.

    The original DataFrame retains exactly the user's columns and values.
    Excel row numbers are tracked separately and never mutate the DataFrame.
    """

    df: pd.DataFrame
    sheet_name: str
    file_source: str
    excel_row_mapping: Dict[int, int]  # DataFrame index -> physical Excel row number


@dataclass(frozen=True)
class SchemaDiff:
    # Legacy fields (preserved for Phase 4 backward compatibility)
    common_columns: List[str]
    only_in_a: List[str]
    only_in_b: List[str]
    key_columns: List[str]
    compare_columns: List[str]

    # New pair-based fields (final selected mappings only)
    exact_name_key_pairs: List[ColumnPair] = field(default_factory=list)
    exact_name_compare_pairs: List[ColumnPair] = field(default_factory=list)
    renamed_key_pairs: List[ColumnPair] = field(default_factory=list)
    renamed_compare_pairs: List[ColumnPair] = field(default_factory=list)
    unmapped_file_a_columns: List[str] = field(default_factory=list)
    unmapped_file_b_columns: List[str] = field(default_factory=list)


@dataclass(frozen=True)
class DuplicateGroup:
    key_values: Tuple[Any, ...]
    key_display: str
    file: str
    occurrences: int
    row_indices: List[int]


@dataclass(frozen=True)
class DuplicateSourceRecord:
    source_file: str
    original_index: int
    excel_row_number: Optional[int]
    key_values: Tuple[Any, ...]
    key_display: str
    group_occurrence_count: int
    all_column_values: Dict[str, Any]


@dataclass(frozen=True)
class MissingKeyRecord:
    source_file: str
    original_index: int
    excel_row_number: Optional[int]
    key_values: Tuple[Any, ...]
    key_display: str
    all_column_values: Dict[str, Any]


@dataclass(frozen=True)
class DataQualityIssue:
    issue_type: str
    source_file: str
    original_index: int
    excel_row_number: Optional[int]
    key_display: str
    all_column_values: Dict[str, Any]
    key_values_a: tuple = ()
    key_values_b: tuple = ()


@dataclass(frozen=True)
class MatchResult:
    matched_pairs: List[Tuple[pd.Series, pd.Series]] = field(default_factory=list)
    only_in_a_unique: List[pd.Series] = field(default_factory=list)
    only_in_b_unique: List[pd.Series] = field(default_factory=list)
    duplicate_groups_a: List[DuplicateGroup] = field(default_factory=list)
    duplicate_groups_b: List[DuplicateGroup] = field(default_factory=list)
    duplicate_source_records_a: List[DuplicateSourceRecord] = field(default_factory=list)
    duplicate_source_records_b: List[DuplicateSourceRecord] = field(default_factory=list)
    missing_key_records_a: List[MissingKeyRecord] = field(default_factory=list)
    missing_key_records_b: List[MissingKeyRecord] = field(default_factory=list)
    schema_diff: Optional[SchemaDiff] = None


@dataclass(frozen=True)
class ComparisonResult:
    matched: List[RowDiff]
    mismatched: List[RowDiff]
    only_in_a: List[RowDiff]
    only_in_b: List[RowDiff]
    data_quality_issues: List[DataQualityIssue]
    schema_diff: SchemaDiff
    summary: Dict[str, Any] = field(default_factory=dict)
