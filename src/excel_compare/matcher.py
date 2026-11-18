"""Composite-key matching.

Joins two DataFrames on configurable key column pairs, classifies rows as
matched / only-in-A / only-in-B, and detects duplicate and missing keys.

Duplicate handling (non-blocking):
- Unique key groups are matched normally.
- Duplicate key groups are NOT automatically matched.
- Every affected source record is preserved with original values.
- Ambiguous groups are reported in Data Quality Issues, not dropped.
- Unambiguous matching continues independently of duplicates.

Missing key values are preserved and reported, never silently discarded.
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from .exceptions import DuplicateKeyError, InvalidKeySelectionError, MappingValidationError
from .models import (
    ColumnPair,
    ComparisonConfig,
    DuplicateGroup,
    DuplicateSourceRecord,
    MatchResult,
    MissingKeyRecord,
    SchemaDiff,
)


def discover_auto_pairs(
    columns_a: List[str],
    columns_b: List[str],
) -> List[ColumnPair]:
    """Return ColumnPair(col, col) for every name present in both files.

    File A order preserved. Exact name match only. No fuzzy matching.
    """
    b_set = set(columns_b)
    pairs: List[ColumnPair] = []
    for col in columns_a:
        if col in b_set:
            pairs.append(ColumnPair(file_a_column=col, file_b_column=col))
    return pairs


def validate_column_mapping(
    key_pairs: List[ColumnPair],
    compare_pairs: List[ColumnPair],
    numeric_pairs: List[ColumnPair],
    date_pairs: List[ColumnPair],
    columns_a: List[str],
    columns_b: List[str],
) -> None:
    """Enforce the column mapping validation contract.

    Raises MappingValidationError on any violation.
    """
    a_cols_set = set(columns_a)
    b_cols_set = set(columns_b)

    # Collect all paired columns per file.
    a_key_cols: List[str] = []
    b_key_cols: List[str] = []
    a_compare_cols: List[str] = []
    b_compare_cols: List[str] = []

    for pair in key_pairs:
        a_key_cols.append(pair.file_a_column)
        b_key_cols.append(pair.file_b_column)
    for pair in compare_pairs:
        a_compare_cols.append(pair.file_a_column)
        b_compare_cols.append(pair.file_b_column)

    # Key mappings cannot be empty.
    if not key_pairs:
        raise MappingValidationError(  # noqa: TRY003
            'At least one matching key pair is required.'
        )

    # Both columns in every pair must exist in their respective files.
    missing_a_key = [c for c in a_key_cols if c not in a_cols_set]
    missing_b_key = [c for c in b_key_cols if c not in b_cols_set]
    if missing_a_key:
        raise MappingValidationError(  # noqa: TRY003
            f'Key column(s) missing from File A: {", ".join(missing_a_key)}'
        )
    if missing_b_key:
        raise MappingValidationError(  # noqa: TRY003
            f'Key column(s) missing from File B: {", ".join(missing_b_key)}'
        )

    missing_a_compare = [c for c in a_compare_cols if c not in a_cols_set]
    missing_b_compare = [c for c in b_compare_cols if c not in b_cols_set]
    if missing_a_compare:
        raise MappingValidationError(  # noqa: TRY003
            f'Compare column(s) missing from File A: {", ".join(missing_a_compare)}'
        )
    if missing_b_compare:
        raise MappingValidationError(  # noqa: TRY003
            f'Compare column(s) missing from File B: {", ".join(missing_b_compare)}'
        )

    # A File A column cannot appear more than once across key and compare mappings.
    all_a_key = a_key_cols
    all_a_compare = a_compare_cols
    dup_a = [c for c in set(all_a_key) if all_a_key.count(c) > 1]
    if dup_a:
        raise MappingValidationError(  # noqa: TRY003
            f'File A column(s) appear more than once in key/compare mappings: '
            f'{", ".join(sorted(dup_a))}'
        )
    dup_a_compare = [c for c in set(all_a_compare) if all_a_compare.count(c) > 1]
    if dup_a_compare:
        raise MappingValidationError(  # noqa: TRY003
            f'File A column(s) appear more than once in compare mappings: '
            f'{", ".join(sorted(dup_a_compare))}'
        )

    # A File B column cannot appear more than once across key and compare mappings.
    all_b_key = b_key_cols
    all_b_compare = b_compare_cols
    dup_b = [c for c in set(all_b_key) if all_b_key.count(c) > 1]
    if dup_b:
        raise MappingValidationError(  # noqa: TRY003
            f'File B column(s) appear more than once in key/compare mappings: '
            f'{", ".join(sorted(dup_b))}'
        )
    dup_b_compare = [c for c in set(all_b_compare) if all_b_compare.count(c) > 1]
    if dup_b_compare:
        raise MappingValidationError(  # noqa: TRY003
            f'File B column(s) appear more than once in compare mappings: '
            f'{", ".join(sorted(dup_b_compare))}'
        )

    # A matching-key column cannot also be a comparison column.
    key_a_set = set(a_key_cols)
    key_b_set = set(b_key_cols)
    compare_a_set = set(a_compare_cols)
    compare_b_set = set(b_compare_cols)
    overlap_a = key_a_set & compare_a_set
    overlap_b = key_b_set & compare_b_set
    if overlap_a:
        raise MappingValidationError(  # noqa: TRY003
            f'File A column(s) used as both matching key and comparison column: '
            f'{", ".join(sorted(overlap_a))}'
        )
    if overlap_b:
        raise MappingValidationError(  # noqa: TRY003
            f'File B column(s) used as both matching key and comparison column: '
            f'{", ".join(sorted(overlap_b))}'
        )

    # Every numeric_pairs entry must belong to compare_pairs.
    compare_set = {(p.file_a_column, p.file_b_column) for p in compare_pairs}
    for pair in numeric_pairs:
        if (pair.file_a_column, pair.file_b_column) not in compare_set:
            raise MappingValidationError(  # noqa: TRY003
                f'Numeric pair ({pair.file_a_column}, {pair.file_b_column}) '
                f'is not in compare_pairs.'
            )

    # Every date_pairs entry must belong to compare_pairs.
    for pair in date_pairs:
        if (pair.file_a_column, pair.file_b_column) not in compare_set:
            raise MappingValidationError(  # noqa: TRY003
                f'Date pair ({pair.file_a_column}, {pair.file_b_column}) is not in compare_pairs.'
            )

    # A pair cannot appear in both numeric_pairs and date_pairs.
    numeric_set = {(p.file_a_column, p.file_b_column) for p in numeric_pairs}
    date_set = {(p.file_a_column, p.file_b_column) for p in date_pairs}
    overlap_nd = numeric_set & date_set
    if overlap_nd:
        raise MappingValidationError(  # noqa: TRY003
            f'Pair(s) appear in both numeric_pairs and date_pairs: '
            f'{", ".join(f"({a}, {b})" for a, b in sorted(overlap_nd))}'
        )


def _config_key_columns(config: ComparisonConfig) -> List[ColumnPair]:
    """Resolve key pairs from config, handling legacy keys field."""
    if config.key_pairs:
        return list(config.key_pairs)
    if config.keys:
        return [ColumnPair(file_a_column=k, file_b_column=k) for k in config.keys]
    return []


def _config_compare_columns(config: ComparisonConfig) -> List[ColumnPair]:
    """Resolve compare pairs from config, handling legacy compare_columns field."""
    if config.compare_pairs:
        return list(config.compare_pairs)
    if config.compare_columns:
        return [ColumnPair(file_a_column=c, file_b_column=c) for c in config.compare_columns]
    return []


def _config_numeric_pairs(config: ComparisonConfig) -> List[ColumnPair]:
    """Resolve numeric pairs from config, handling legacy numeric_columns field."""
    if config.numeric_pairs:
        return list(config.numeric_pairs)
    if config.numeric_columns:
        return [ColumnPair(file_a_column=c, file_b_column=c) for c in config.numeric_columns]
    return []


def _config_date_pairs(config: ComparisonConfig) -> List[ColumnPair]:
    """Resolve date pairs from config, handling legacy date_columns field."""
    if config.date_pairs:
        return list(config.date_pairs)
    if config.date_columns:
        return [ColumnPair(file_a_column=c, file_b_column=c) for c in config.date_columns]
    return []


def _build_key_tuples(
    df: pd.DataFrame,
    columns: List[str],
    case_sensitive: bool,
    normalize_whitespace: bool,
) -> List[Tuple[Any, ...]]:
    """Build normalized composite key tuples for each row.

    Normalization is applied only for matching identity. Original values
    in the DataFrame are never modified.
    """
    tuples: List[Tuple[Any, ...]] = []
    for _, row in df.iterrows():
        key_vals: List[Any] = []
        for k in columns:
            v = row[k]
            # Normalize pandas NaN to None for consistent missing handling.
            if isinstance(v, float) and pd.isna(v):
                v = None
            if isinstance(v, str):
                if normalize_whitespace:
                    v = v.strip()
                if not case_sensitive:
                    v = v.lower()
            key_vals.append(v)
        tuples.append(tuple(key_vals))
    return tuples


def _key_is_missing(key_tuple: Tuple[Any, ...]) -> bool:
    """Return True if any element of the key tuple is missing or empty."""
    for value in key_tuple:
        if value is None:
            return True
        if isinstance(value, float) and pd.isna(value):
            return True
        if isinstance(value, str) and value.strip() == '':
            return True
    return False


def detect_duplicates(
    key_tuples: List[Tuple[Any, ...]],
) -> Dict[Tuple[Any, ...], List[int]]:
    """Return mapping of duplicate composite key -> list of row indices.

    Uses collision-safe tuples, never delimiter-joined strings.
    """
    groups: Dict[Tuple[Any, ...], List[int]] = {}
    for idx, key_tuple in enumerate(key_tuples):
        groups.setdefault(key_tuple, []).append(idx)
    return {key: indices for key, indices in groups.items() if len(indices) > 1}


def _key_display(key_values: Tuple[Any, ...], key_pairs: List[ColumnPair]) -> str:
    """Build human-readable key display string (presentation only)."""
    parts = [f'{pair.file_a_column}={v!r}' for pair, v in zip(key_pairs, key_values, strict=False)]
    return ', '.join(parts)


def build_schema_diff(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    key_pairs: Optional[List[ColumnPair]] = None,
    keys: Optional[List[str]] = None,
    compare_pairs: Optional[List[ColumnPair]] = None,
) -> SchemaDiff:
    """Compute schema differences with deterministic column ordering.

    File A's original column order is preserved. Sets are used only for
    membership tests, never for output ordering.
    """
    # Resolve legacy keys -> key_pairs
    if key_pairs is None and keys is not None:
        key_pairs = [ColumnPair(file_a_column=k, file_b_column=k) for k in keys]
    if key_pairs is None:
        key_pairs = []
    # Resolve legacy compare_columns -> compare_pairs
    if compare_pairs is None:
        compare_pairs = []
    b_cols = set(df_b.columns)
    a_cols = set(df_a.columns)

    common_columns = [c for c in df_a.columns if c in b_cols]
    only_in_a = [c for c in df_a.columns if c not in b_cols]
    only_in_b = [c for c in df_b.columns if c not in a_cols]

    key_a_cols = [p.file_a_column for p in key_pairs]
    key_b_cols = [p.file_b_column for p in key_pairs]
    compare_a_cols = [p.file_a_column for p in compare_pairs]
    compare_b_cols = [p.file_b_column for p in compare_pairs]

    key_columns = [k for k in key_a_cols if k in common_columns]
    keys_set = set(key_a_cols)
    compare_columns = [c for c in common_columns if c not in keys_set]

    # Classify selected pairs by name-equality.
    exact_name_key_pairs = [p for p in key_pairs if p.file_a_column == p.file_b_column]
    renamed_key_pairs = [p for p in key_pairs if p.file_a_column != p.file_b_column]
    exact_name_compare_pairs = [p for p in compare_pairs if p.file_a_column == p.file_b_column]
    renamed_compare_pairs = [p for p in compare_pairs if p.file_a_column != p.file_b_column]

    # Unmapped columns: present in a file but not in any pair (key or compare).
    paired_a_cols = set(key_a_cols) | set(compare_a_cols)
    paired_b_cols = set(key_b_cols) | set(compare_b_cols)
    unmapped_a = [c for c in df_a.columns if c not in paired_a_cols]
    unmapped_b = [c for c in df_b.columns if c not in paired_b_cols]

    return SchemaDiff(
        common_columns=common_columns,
        only_in_a=only_in_a,
        only_in_b=only_in_b,
        key_columns=key_columns,
        compare_columns=compare_columns,
        exact_name_key_pairs=exact_name_key_pairs,
        exact_name_compare_pairs=exact_name_compare_pairs,
        renamed_key_pairs=renamed_key_pairs,
        renamed_compare_pairs=renamed_compare_pairs,
        unmapped_file_a_columns=unmapped_a,
        unmapped_file_b_columns=unmapped_b,
    )


def match_on_keys(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    key_pairs: Optional[List[ColumnPair]] = None,
    keys: Optional[List[str]] = None,
    config: Optional[ComparisonConfig] = None,
    excel_row_mapping_a: Optional[Dict[int, int]] = None,
    excel_row_mapping_b: Optional[Dict[int, int]] = None,
) -> MatchResult:
    """Classify rows with non-blocking duplicate handling.

    Parameters
    ----------
    df_a, df_b : pd.DataFrame
        Original DataFrames (unchanged).
    key_pairs : list[ColumnPair] | None
        Composite matching key column pairs. If None, resolved from config.
    config : ComparisonConfig | None
        Comparison settings. If None, defaults are used.
    excel_row_mapping_a, excel_row_mapping_b : dict[int, int] | None
        DataFrame index -> physical Excel row number mapping.

    Returns
    -------
    MatchResult
        All source rows accounted for exactly once across categories.

    Raises
    ------
    InvalidKeySelectionError
        When a key column is missing from either DataFrame.
    DuplicateKeyError
        When duplicate_policy is 'all' (unsupported).
    MappingValidationError
        When column mapping validation fails.
    """
    if config is None:
        config = ComparisonConfig()

    # Resolve key pairs: explicit argument takes precedence, then legacy keys,
    # then config.key_pairs.
    if key_pairs is not None:
        resolved_key_pairs = list(key_pairs)
    elif keys is not None:
        resolved_key_pairs = [ColumnPair(file_a_column=k, file_b_column=k) for k in keys]
    elif config.key_pairs:
        resolved_key_pairs = list(config.key_pairs)
    elif config.keys:
        resolved_key_pairs = [ColumnPair(file_a_column=k, file_b_column=k) for k in config.keys]
    else:
        resolved_key_pairs = []

    # Resolve compare pairs from config.
    resolved_compare_pairs = _config_compare_columns(config)
    resolved_numeric_pairs = _config_numeric_pairs(config)
    resolved_date_pairs = _config_date_pairs(config)

    key_a_cols = [p.file_a_column for p in resolved_key_pairs]
    key_b_cols = [p.file_b_column for p in resolved_key_pairs]

    # Validate key columns exist in both files first (InvalidKeySelectionError
    # takes precedence over MappingValidationError for missing columns).
    missing_a = [k for k in key_a_cols if k not in df_a.columns]
    missing_b = [k for k in key_b_cols if k not in df_b.columns]
    if missing_a:
        raise InvalidKeySelectionError(  # noqa: TRY003
            f'Key column(s) missing from File A: {", ".join(missing_a)}'
        )
    if missing_b:
        raise InvalidKeySelectionError(  # noqa: TRY003
            f'Key column(s) missing from File B: {", ".join(missing_b)}'
        )

    if config.duplicate_policy == 'all':
        raise DuplicateKeyError(  # noqa: TRY003
            'Duplicate key policy "all" is not supported.'
        )

    # Validate full mapping contract (duplicates, key-compare overlap,
    # numeric/date membership). Missing columns already checked above.
    validate_column_mapping(
        resolved_key_pairs,
        resolved_compare_pairs,
        resolved_numeric_pairs,
        resolved_date_pairs,
        list(df_a.columns),
        list(df_b.columns),
    )

    case_sensitive = config.case_sensitive
    normalize_ws = config.normalize_whitespace

    # Build normalized key tuples for matching identity only.
    key_tuples_a = _build_key_tuples(df_a, key_a_cols, case_sensitive, normalize_ws)
    key_tuples_b = _build_key_tuples(df_b, key_b_cols, case_sensitive, normalize_ws)

    # Detect missing/empty key values.
    missing_key_records_a: List[MissingKeyRecord] = []
    missing_key_records_b: List[MissingKeyRecord] = []
    valid_row_indices_a: List[int] = []
    valid_row_indices_b: List[int] = []

    for idx, kt in enumerate(key_tuples_a):
        if _key_is_missing(kt):
            missing_key_records_a.append(
                MissingKeyRecord(
                    source_file='A',
                    original_index=int(df_a.index[idx]),
                    excel_row_number=(
                        excel_row_mapping_a.get(int(df_a.index[idx]))
                        if excel_row_mapping_a
                        else None
                    ),
                    key_values=kt,
                    key_display=_key_display(kt, resolved_key_pairs),
                    all_column_values=df_a.iloc[idx].to_dict(),
                )
            )
        else:
            valid_row_indices_a.append(idx)

    for idx, kt in enumerate(key_tuples_b):
        if _key_is_missing(kt):
            missing_key_records_b.append(
                MissingKeyRecord(
                    source_file='B',
                    original_index=int(df_b.index[idx]),
                    excel_row_number=(
                        excel_row_mapping_b.get(int(df_b.index[idx]))
                        if excel_row_mapping_b
                        else None
                    ),
                    key_values=kt,
                    key_display=_key_display(kt, resolved_key_pairs),
                    all_column_values=df_b.iloc[idx].to_dict(),
                )
            )
        else:
            valid_row_indices_b.append(idx)

    # Detect duplicates on valid (non-missing) key tuples only.
    valid_key_tuples_a = [key_tuples_a[i] for i in valid_row_indices_a]
    valid_key_tuples_b = [key_tuples_b[i] for i in valid_row_indices_b]

    # Map from valid-row position to original df index.
    valid_to_orig_a = {i: valid_row_indices_a[i] for i in range(len(valid_row_indices_a))}
    valid_to_orig_b = {i: valid_row_indices_b[i] for i in range(len(valid_row_indices_b))}

    duplicates_a = detect_duplicates(valid_key_tuples_a)
    duplicates_b = detect_duplicates(valid_key_tuples_b)

    # Combined ambiguous keys: duplicated in either file.
    ambiguous_keys: set = set(duplicates_a.keys()) | set(duplicates_b.keys())

    # Build lookup: key_tuple -> list of valid positions (for all valid rows).
    key_to_indices_a: Dict[Tuple[Any, ...], List[int]] = {}
    for pos in valid_row_indices_a:
        kt = key_tuples_a[pos]
        key_to_indices_a.setdefault(kt, []).append(pos)

    key_to_indices_b: Dict[Tuple[Any, ...], List[int]] = {}
    for pos in valid_row_indices_b:
        kt = key_tuples_b[pos]
        key_to_indices_b.setdefault(kt, []).append(pos)

    # Build duplicate groups and source records for ambiguous keys.
    duplicate_groups_a: List[DuplicateGroup] = []
    duplicate_source_records_a: List[DuplicateSourceRecord] = []
    duplicate_groups_b: List[DuplicateGroup] = []
    duplicate_source_records_b: List[DuplicateSourceRecord] = []

    for kt in ambiguous_keys:
        # File A records with this ambiguous key.
        if kt in key_to_indices_a:
            orig_indices_a = [valid_to_orig_a[p] for p in key_to_indices_a[kt]]
            duplicate_groups_a.append(
                DuplicateGroup(
                    key_values=kt,
                    key_display=_key_display(kt, resolved_key_pairs),
                    file='A',
                    occurrences=len(orig_indices_a),
                    row_indices=orig_indices_a,
                )
            )
            for orig_idx in key_to_indices_a[kt]:
                row = df_a.iloc[orig_idx]
                duplicate_source_records_a.append(
                    DuplicateSourceRecord(
                        source_file='A',
                        original_index=int(df_a.index[orig_idx]),
                        excel_row_number=(
                            excel_row_mapping_a.get(int(df_a.index[orig_idx]))
                            if excel_row_mapping_a
                            else None
                        ),
                        key_values=kt,
                        key_display=_key_display(kt, resolved_key_pairs),
                        group_occurrence_count=len(key_to_indices_a[kt]),
                        all_column_values=row.to_dict(),
                    )
                )

        # File B records with this ambiguous key.
        if kt in key_to_indices_b:
            orig_indices_b = [valid_to_orig_b[p] for p in key_to_indices_b[kt]]
            duplicate_groups_b.append(
                DuplicateGroup(
                    key_values=kt,
                    key_display=_key_display(kt, resolved_key_pairs),
                    file='B',
                    occurrences=len(orig_indices_b),
                    row_indices=orig_indices_b,
                )
            )
            for orig_idx in key_to_indices_b[kt]:
                row = df_b.iloc[orig_idx]
                duplicate_source_records_b.append(
                    DuplicateSourceRecord(
                        source_file='B',
                        original_index=int(df_b.index[orig_idx]),
                        excel_row_number=(
                            excel_row_mapping_b.get(int(df_b.index[orig_idx]))
                            if excel_row_mapping_b
                            else None
                        ),
                        key_values=kt,
                        key_display=_key_display(kt, resolved_key_pairs),
                        group_occurrence_count=len(key_to_indices_b[kt]),
                        all_column_values=row.to_dict(),
                    )
                )

    # Classify unambiguous keys: keys appearing exactly once in both files,
    # and not duplicated on either side.
    all_keys_a = set(key_to_indices_a.keys())
    all_keys_b = set(key_to_indices_b.keys())

    unambiguous_keys = (all_keys_a & all_keys_b) - ambiguous_keys

    matched_pairs: List[Tuple[pd.Series, pd.Series]] = []
    only_in_a_unique: List[pd.Series] = []
    only_in_b_unique: List[pd.Series] = []

    # Matched pairs: keys appearing exactly once in both files.
    for kt in unambiguous_keys:
        idx_a = key_to_indices_a[kt][0]
        idx_b = key_to_indices_b[kt][0]
        matched_pairs.append((df_a.iloc[idx_a], df_b.iloc[idx_b]))

    # Only in A: keys in A but not in B, and not ambiguous.
    only_a_keys = (all_keys_a - all_keys_b) - ambiguous_keys
    for kt in only_a_keys:
        for pos in key_to_indices_a[kt]:
            orig_idx = valid_to_orig_a[pos]
            only_in_a_unique.append(df_a.iloc[orig_idx])

    # Only in B: keys in B but not in A, and not ambiguous.
    only_b_keys = (all_keys_b - all_keys_a) - ambiguous_keys
    for kt in only_b_keys:
        for pos in key_to_indices_b[kt]:
            orig_idx = valid_to_orig_b[pos]
            only_in_b_unique.append(df_b.iloc[orig_idx])

    # Build schema diff.
    schema_diff = build_schema_diff(df_a, df_b, resolved_key_pairs, resolved_compare_pairs)

    return MatchResult(
        matched_pairs=matched_pairs,
        only_in_a_unique=only_in_a_unique,
        only_in_b_unique=only_in_b_unique,
        duplicate_groups_a=duplicate_groups_a,
        duplicate_groups_b=duplicate_groups_b,
        duplicate_source_records_a=duplicate_source_records_a,
        duplicate_source_records_b=duplicate_source_records_b,
        missing_key_records_a=missing_key_records_a,
        missing_key_records_b=missing_key_records_b,
        schema_diff=schema_diff,
    )
