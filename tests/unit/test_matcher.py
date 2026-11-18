"""Unit tests for the matcher module (Phase 4)."""

from __future__ import annotations

from datetime import date, datetime

import pandas as pd
import pytest

from excel_compare.exceptions import (
    DuplicateKeyError,
    InvalidKeySelectionError,
    MappingValidationError,
)
from excel_compare.matcher import (
    build_schema_diff,
    detect_duplicates,
    discover_auto_pairs,
    match_on_keys,
)
from excel_compare.models import ColumnPair, ComparisonConfig


def _df(rows: list[dict], index: list | None = None) -> pd.DataFrame:
    return pd.DataFrame(rows, index=index)


def _config(
    keys=None,
    case_sensitive: bool = False,
    normalize_whitespace: bool = True,
    duplicate_policy: str = 'error',
) -> ComparisonConfig:
    return ComparisonConfig(
        keys=keys or [],
        case_sensitive=case_sensitive,
        normalize_whitespace=normalize_whitespace,
        duplicate_policy=duplicate_policy,
    )


def test_single_key_matching():
    df_a = _df(
        [
            {'id': '1', 'name': 'Alice', 'amount': '100'},
            {'id': '2', 'name': 'Bob', 'amount': '200'},
        ]
    )
    df_b = _df(
        [
            {'id': '1', 'name': 'Alice', 'amount': '100'},
            {'id': '3', 'name': 'Carol', 'amount': '300'},
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert len(result.matched_pairs) == 1
    assert len(result.only_in_a_unique) == 1
    assert len(result.only_in_b_unique) == 1
    assert result.only_in_a_unique[0]['id'] == '2'
    assert result.only_in_b_unique[0]['id'] == '3'


def test_composite_key_matching():
    df_a = _df(
        [
            {'id': '1', 'date': '2024-01-01', 'name': 'Alice'},
            {'id': '1', 'date': '2024-01-02', 'name': 'Bob'},
        ]
    )
    df_b = _df(
        [
            {'id': '1', 'date': '2024-01-01', 'name': 'Alice'},
            {'id': '2', 'date': '2024-01-01', 'name': 'Carol'},
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['id', 'date'], config=_config(keys=['id', 'date']))
    assert len(result.matched_pairs) == 1
    assert result.matched_pairs[0][0]['name'] == 'Alice'
    assert len(result.only_in_a_unique) == 1
    assert result.only_in_a_unique[0]['name'] == 'Bob'
    assert len(result.only_in_b_unique) == 1
    assert result.only_in_b_unique[0]['name'] == 'Carol'


def test_composite_key_collision_delimiter_characters():
    """Tuple-based keys must not collide when values contain '|'."""
    df_a = _df(
        [
            {'k1': 'A|B', 'k2': 'C', 'val': 'x'},
            {'k1': 'A', 'k2': 'B|C', 'val': 'y'},
        ]
    )
    df_b = _df(
        [
            {'k1': 'A|B', 'k2': 'C', 'val': 'x'},
            {'k1': 'A', 'k2': 'B|C', 'val': 'y'},
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['k1', 'k2'], config=_config(keys=['k1', 'k2']))
    assert len(result.matched_pairs) == 2
    assert len(result.only_in_a_unique) == 0
    assert len(result.only_in_b_unique) == 0


def test_missing_key_column_raises():
    df_a = _df([{'id': '1', 'name': 'Alice'}])
    df_b = _df([{'name': 'Alice'}])
    with pytest.raises(InvalidKeySelectionError, match='missing from File B'):
        match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))


def test_empty_dataframes_with_keys():
    df_a = pd.DataFrame(columns=['id', 'name'])
    df_b = pd.DataFrame(columns=['id', 'name'])
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert result.matched_pairs == []
    assert result.only_in_a_unique == []
    assert result.only_in_b_unique == []
    assert result.duplicate_groups_a == []
    assert result.duplicate_groups_b == []
    assert result.duplicate_source_records_a == []
    assert result.duplicate_source_records_b == []
    assert result.missing_key_records_a == []
    assert result.missing_key_records_b == []


def test_missing_key_values_preserved():
    df_a = _df(
        [
            {'id': '', 'name': 'Alice'},
            {'id': '2', 'name': 'Bob'},
        ]
    )
    df_b = _df(
        [
            {'id': '2', 'name': 'Bob'},
            {'id': '3', 'name': 'Carol'},
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert len(result.missing_key_records_a) == 1
    assert result.missing_key_records_a[0].key_values == ('',)
    assert result.missing_key_records_a[0].all_column_values == {'id': '', 'name': 'Alice'}
    # Missing-key record must NOT appear in only_in_a_unique
    assert all(r['id'] != '' for r in result.only_in_a_unique)
    # Valid record still matched
    assert len(result.matched_pairs) == 1
    assert result.matched_pairs[0][0]['id'] == '2'


def test_none_key_values_preserved():
    df_a = _df(
        [
            {'id': None, 'name': 'Alice'},
            {'id': '2', 'name': 'Bob'},
        ]
    )
    df_b = _df(
        [
            {'id': '2', 'name': 'Bob'},
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert len(result.missing_key_records_a) == 1
    assert result.missing_key_records_a[0].key_values == (None,)


def test_duplicate_keys_not_auto_matched():
    """Duplicate key group must not be automatically matched."""
    df_a = _df(
        [
            {'id': '1', 'name': 'Alice'},
            {'id': '1', 'name': 'Alice2'},
        ]
    )
    df_b = _df(
        [
            {'id': '1', 'name': 'Alice'},
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    # None of the three records is matched
    assert len(result.matched_pairs) == 0
    # All three are classified as ambiguous
    assert len(result.duplicate_source_records_a) == 2
    assert len(result.duplicate_source_records_b) == 1
    # The unique record in File B must NOT appear in only_in_b_unique
    assert len(result.only_in_b_unique) == 0
    # Duplicate groups detected
    assert len(result.duplicate_groups_a) == 1
    assert result.duplicate_groups_a[0].occurrences == 2
    assert len(result.duplicate_groups_b) == 1
    assert result.duplicate_groups_b[0].occurrences == 1


def test_duplicate_keys_reverse_situation():
    """Reverse: File A unique, File B duplicated."""
    df_a = _df(
        [
            {'id': '1', 'name': 'Alice'},
        ]
    )
    df_b = _df(
        [
            {'id': '1', 'name': 'Alice'},
            {'id': '1', 'name': 'Alice2'},
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert len(result.matched_pairs) == 0
    assert len(result.duplicate_source_records_a) == 1
    assert len(result.duplicate_source_records_b) == 2
    assert len(result.only_in_a_unique) == 0
    assert len(result.only_in_b_unique) == 0


def test_duplicate_policy_error_non_blocking():
    """'error' policy reports duplicates but does not block unambiguous matching."""
    df_a = _df(
        [
            {'id': '1', 'name': 'Alice'},
            {'id': '1', 'name': 'Alice2'},
            {'id': '2', 'name': 'Bob'},
        ]
    )
    df_b = _df(
        [
            {'id': '1', 'name': 'Alice'},
            {'id': '2', 'name': 'Bob'},
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    # Unambiguous key '2' is matched
    assert len(result.matched_pairs) == 1
    assert result.matched_pairs[0][0]['id'] == '2'
    # Duplicate key '1' is reported, not matched
    assert len(result.duplicate_source_records_a) == 2
    assert len(result.duplicate_source_records_b) == 1


def test_duplicate_policy_all_unsupported():
    df_a = _df(
        [
            {'id': '1', 'name': 'Alice'},
            {'id': '1', 'name': 'Alice2'},
        ]
    )
    df_b = _df([{'id': '1', 'name': 'Alice'}])
    with pytest.raises(DuplicateKeyError, match='not supported'):
        match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id'], duplicate_policy='all'))


def test_detect_duplicates_returns_empty_for_no_duplicates():
    key_tuples = [('1',), ('2',), ('3',)]
    assert detect_duplicates(key_tuples) == {}


def test_detect_duplicates_returns_indices():
    key_tuples = [('1',), ('2',), ('1',)]
    dups = detect_duplicates(key_tuples)
    assert ('1',) in dups
    assert dups[('1',)] == [0, 2]


def test_original_indices_preserved():
    df_a = _df(
        [
            {'id': '1', 'name': 'Alice'},
            {'id': '2', 'name': 'Bob'},
        ],
        index=[10, 20],
    )
    df_b = _df(
        [
            {'id': '2', 'name': 'Bob'},
            {'id': '3', 'name': 'Carol'},
        ],
        index=[30, 40],
    )
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert result.matched_pairs[0][0].name == 20
    assert result.matched_pairs[0][1].name == 30


def test_schema_diff_deterministic_ordering():
    df_a = _df([{'z': '1', 'a': '2', 'm': '3'}])
    df_b = _df([{'z': '1', 'a': '2', 'x': '4'}])
    schema = build_schema_diff(df_a, df_b, keys=['z'])
    # File A order preserved
    assert schema.common_columns == ['z', 'a']
    assert schema.only_in_a == ['m']
    assert schema.only_in_b == ['x']
    assert schema.key_columns == ['z']
    assert schema.compare_columns == ['a']


def test_duplicate_records_preserve_original_values():
    df_a = _df(
        [
            {'id': '1', 'name': ' Alice ', 'amount': '100'},
            {'id': '1', 'name': 'ALICE', 'amount': '200'},
        ]
    )
    df_b = _df(
        [
            {'id': '1', 'name': 'Alice', 'amount': '100'},
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert len(result.duplicate_source_records_a) == 2
    # Original whitespace and capitalization preserved
    assert result.duplicate_source_records_a[0].all_column_values['name'] == ' Alice '
    assert result.duplicate_source_records_a[1].all_column_values['name'] == 'ALICE'
    assert result.duplicate_source_records_a[0].group_occurrence_count == 2
    assert result.duplicate_source_records_b[0].group_occurrence_count == 1


def test_every_source_row_accounted_for():
    """Each source row appears in exactly one category."""
    df_a = _df(
        [
            {'id': '1', 'name': 'Alice'},
            {'id': '1', 'name': 'Alice2'},  # duplicate
            {'id': '3', 'name': 'Carol'},  # only in A
            {'id': '', 'name': 'Missing'},  # missing key
        ]
    )
    df_b = _df(
        [
            {'id': '1', 'name': 'Alice'},  # ambiguous with A's dups
            {'id': '2', 'name': 'Bob'},  # only in B
            {'id': '', 'name': 'Missing2'},  # missing key
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))

    # File A accounting: 4 input rows
    a_accounted = (
        len(result.matched_pairs)
        + len(result.only_in_a_unique)
        + len(result.duplicate_source_records_a)
        + len(result.missing_key_records_a)
    )
    assert a_accounted == 4, f'File A: {a_accounted} != 4'

    # File B accounting: 3 input rows
    b_accounted = (
        len(result.matched_pairs)
        + len(result.only_in_b_unique)
        + len(result.duplicate_source_records_b)
        + len(result.missing_key_records_b)
    )
    assert b_accounted == 3, f'File B: {b_accounted} != 3'

    # No overlap: duplicate records not in matched or only-in
    matched_ids_a = {p[0]['id'] for p in result.matched_pairs}
    only_a_ids = {r['id'] for r in result.only_in_a_unique}
    dup_ids_a = {r.all_column_values['id'] for r in result.duplicate_source_records_a}
    dup_ids_b = {r.all_column_values['id'] for r in result.duplicate_source_records_b}

    # Matched and only-in should not overlap with duplicates
    for mid in matched_ids_a:
        assert mid not in dup_ids_a
        assert mid not in dup_ids_b
    for oid in only_a_ids:
        assert oid not in dup_ids_a
        assert oid not in dup_ids_b


def test_every_source_row_accounted_for_renamed():
    """Each source row accounted for exactly once with renamed keys."""
    df_a = _df(
        [
            {'Site ID': 'S1', 'Pole': 'P1', 'name': 'Alice'},
            {'Site ID': 'S1', 'Pole': 'P1', 'name': 'Alice2'},
            {'Site ID': 'S3', 'Pole': 'P3', 'name': 'Carol'},
            {'Site ID': '', 'Pole': 'P4', 'name': 'Missing'},
        ]
    )
    df_b = _df(
        [
            {'Site Code': 'S1', 'Pole No': 'P1', 'name': 'Alice'},
            {'Site Code': 'S2', 'Pole No': 'P2', 'name': 'Bob'},
            {'Site Code': '', 'Pole No': 'P4', 'name': 'Missing2'},
        ]
    )
    key_pairs = [_pair('Site ID', 'Site Code'), _pair('Pole', 'Pole No')]
    config = ComparisonConfig(keys=[], key_pairs=key_pairs)
    result = match_on_keys(df_a, df_b, key_pairs=key_pairs, config=config)

    a_accounted = (
        len(result.matched_pairs)
        + len(result.only_in_a_unique)
        + len(result.duplicate_source_records_a)
        + len(result.missing_key_records_a)
    )
    assert a_accounted == 4

    b_accounted = (
        len(result.matched_pairs)
        + len(result.only_in_b_unique)
        + len(result.duplicate_source_records_b)
        + len(result.missing_key_records_b)
    )
    assert b_accounted == 3


def _pair(a_col: str, b_col: str) -> ColumnPair:
    return ColumnPair(file_a_column=a_col, file_b_column=b_col)


def test_renamed_key_pair_matching():
    """Composite key with different column names on each side."""
    df_a = _df(
        [
            {'Site ID': 'S1', 'Pole': 'P1', 'name': 'Alice'},
            {'Site ID': 'S2', 'Pole': 'P2', 'name': 'Bob'},
        ]
    )
    df_b = _df(
        [
            {'Site Code': 'S1', 'Pole No': 'P1', 'name': 'Alice'},
            {'Site Code': 'S3', 'Pole No': 'P3', 'name': 'Carol'},
        ]
    )
    key_pairs = [_pair('Site ID', 'Site Code'), _pair('Pole', 'Pole No')]
    config = ComparisonConfig(keys=[], key_pairs=key_pairs)
    result = match_on_keys(df_a, df_b, key_pairs=key_pairs, config=config)
    assert len(result.matched_pairs) == 1
    assert result.matched_pairs[0][0]['Site ID'] == 'S1'
    assert len(result.only_in_a_unique) == 1
    assert result.only_in_a_unique[0]['Site ID'] == 'S2'
    assert len(result.only_in_b_unique) == 1
    assert result.only_in_b_unique[0]['Site Code'] == 'S3'


def test_duplicate_keys_renamed_not_auto_matched():
    """Duplicate renamed key group must not be automatically matched."""
    df_a = _df(
        [
            {'Site ID': 'S1', 'Pole': 'P1'},
            {'Site ID': 'S1', 'Pole': 'P1'},
        ]
    )
    df_b = _df(
        [
            {'Site Code': 'S1', 'Pole No': 'P1'},
        ]
    )
    key_pairs = [_pair('Site ID', 'Site Code'), _pair('Pole', 'Pole No')]
    config = ComparisonConfig(keys=[], key_pairs=key_pairs)
    result = match_on_keys(df_a, df_b, key_pairs=key_pairs, config=config)
    assert len(result.matched_pairs) == 0
    assert len(result.duplicate_source_records_a) == 2
    assert len(result.duplicate_source_records_b) == 1
    assert len(result.only_in_b_unique) == 0


def test_missing_key_values_renamed_preserved():
    """Missing key values with renamed pairs are preserved."""
    df_a = _df(
        [
            {'Site ID': '', 'Pole': 'P1'},
            {'Site ID': 'S2', 'Pole': 'P2'},
        ]
    )
    df_b = _df(
        [
            {'Site Code': 'S2', 'Pole No': 'P2'},
        ]
    )
    key_pairs = [_pair('Site ID', 'Site Code'), _pair('Pole', 'Pole No')]
    config = ComparisonConfig(keys=[], key_pairs=key_pairs)
    result = match_on_keys(df_a, df_b, key_pairs=key_pairs, config=config)
    assert len(result.missing_key_records_a) == 1
    assert result.missing_key_records_a[0].key_values == ('', 'p1')
    assert result.missing_key_records_a[0].all_column_values == {'Site ID': '', 'Pole': 'P1'}
    assert len(result.matched_pairs) == 1
    assert result.matched_pairs[0][0]['Site ID'] == 'S2'


def test_cross_file_ambiguity_renamed_keys():
    """File A unique key, File B duplicated key → all ambiguous."""
    df_a = _df(
        [
            {'Site ID': 'S1', 'Pole': 'P1'},
        ]
    )
    df_b = _df(
        [
            {'Site Code': 'S1', 'Pole No': 'P1'},
            {'Site Code': 'S1', 'Pole No': 'P1'},
        ]
    )
    key_pairs = [_pair('Site ID', 'Site Code'), _pair('Pole', 'Pole No')]
    config = ComparisonConfig(keys=[], key_pairs=key_pairs)
    result = match_on_keys(df_a, df_b, key_pairs=key_pairs, config=config)
    assert len(result.matched_pairs) == 0
    assert len(result.duplicate_source_records_a) == 1
    assert len(result.duplicate_source_records_b) == 2
    assert len(result.only_in_a_unique) == 0
    assert len(result.only_in_b_unique) == 0


def test_duplicate_policy_error_non_blocking_renamed():
    """'error' policy reports duplicates but continues unambiguous matching."""
    df_a = _df(
        [
            {'Site ID': 'S1', 'Pole': 'P1'},
            {'Site ID': 'S1', 'Pole': 'P1'},
            {'Site ID': 'S2', 'Pole': 'P2'},
        ]
    )
    df_b = _df(
        [
            {'Site Code': 'S1', 'Pole No': 'P1'},
            {'Site Code': 'S2', 'Pole No': 'P2'},
        ]
    )
    key_pairs = [_pair('Site ID', 'Site Code'), _pair('Pole', 'Pole No')]
    config = ComparisonConfig(keys=[], key_pairs=key_pairs)
    result = match_on_keys(df_a, df_b, key_pairs=key_pairs, config=config)
    assert len(result.matched_pairs) == 1
    assert result.matched_pairs[0][0]['Site ID'] == 'S2'
    assert len(result.duplicate_source_records_a) == 2
    assert len(result.duplicate_source_records_b) == 1


def test_mapping_missing_column_file_a_raises():
    """Missing key column in File A raises InvalidKeySelectionError."""
    df_a = _df([{'name': 'Alice'}])
    df_b = _df([{'Site Code': 'S1'}])
    key_pairs = [_pair('Site ID', 'Site Code')]
    config = ComparisonConfig(keys=[], key_pairs=key_pairs)
    with pytest.raises(InvalidKeySelectionError, match='missing from File A'):
        match_on_keys(df_a, df_b, key_pairs=key_pairs, config=config)


def test_mapping_missing_column_file_b_raises():
    """Missing key column in File B raises InvalidKeySelectionError."""
    df_a = _df([{'Site ID': 'S1'}])
    df_b = _df([{'name': 'Alice'}])
    key_pairs = [_pair('Site ID', 'Site Code')]
    config = ComparisonConfig(keys=[], key_pairs=key_pairs)
    with pytest.raises(InvalidKeySelectionError, match='missing from File B'):
        match_on_keys(df_a, df_b, key_pairs=key_pairs, config=config)


def test_mapping_duplicate_file_a_column_raises():
    """Same File A column in key and compare raises error."""
    df_a = _df([{'id': '1', 'name': 'A'}])
    df_b = _df([{'id': '1', 'name': 'A'}])
    key_pairs = [_pair('id', 'id')]
    compare_pairs = [_pair('id', 'id')]
    config = ComparisonConfig(keys=[], key_pairs=key_pairs, compare_pairs=compare_pairs)
    with pytest.raises(MappingValidationError, match='both matching key and comparison'):
        match_on_keys(df_a, df_b, key_pairs=key_pairs, config=config)


def test_mapping_empty_keys_raises():
    """Empty key pairs raises error."""
    df_a = _df([{'id': '1'}])
    df_b = _df([{'id': '1'}])
    config = ComparisonConfig(keys=[], key_pairs=[], compare_pairs=[])
    with pytest.raises(MappingValidationError, match='At least one matching key pair'):
        match_on_keys(df_a, df_b, key_pairs=[], config=config)


def test_discover_auto_pairs_exact_name_only():
    """Auto-discovery only pairs identical names."""
    pairs = discover_auto_pairs(['Site ID', 'Pole', 'Name'], ['Site Code', 'Pole', 'Name'])
    assert len(pairs) == 2
    assert ColumnPair(file_a_column='Pole', file_b_column='Pole') in pairs
    assert ColumnPair(file_a_column='Name', file_b_column='Name') in pairs


# ---------------------------------------------------------------------------
# Matching policy regression tests (Issue 1)
# Numeric vs text key matching preserves existing behaviour via str
# canonicalization of key identities.
# ---------------------------------------------------------------------------


def test_numeric_key_matches_text_key():
    """Numeric 123 matches text '123' (existing behaviour preserved)."""
    df_a = _df([{'id': 123, 'name': 'Alice'}])
    df_b = _df([{'id': '123', 'name': 'Alice'}])
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert len(result.matched_pairs) == 1
    assert len(result.only_in_a_unique) == 0
    assert len(result.only_in_b_unique) == 0


def test_numeric_key_vs_leading_zero_text_key_not_match():
    """Numeric 123 does NOT match text '00123'."""
    df_a = _df([{'id': 123, 'name': 'Alice'}])
    df_b = _df([{'id': '00123', 'name': 'Alice'}])
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert len(result.matched_pairs) == 0
    assert len(result.only_in_a_unique) == 1
    assert len(result.only_in_b_unique) == 1


def test_numeric_key_matches_numeric_key():
    """Numeric 123 matches numeric 123."""
    df_a = _df([{'id': 123, 'name': 'Alice'}])
    df_b = _df([{'id': 123, 'name': 'Alice'}])
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert len(result.matched_pairs) == 1


def test_numeric_key_vs_float_key_not_match():
    """Numeric 123 vs numeric 123.0 do not match (str(123) != str(123.0))."""
    df_a = _df([{'id': 123, 'name': 'Alice'}])
    df_b = _df([{'id': 123.0, 'name': 'Alice'}])
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert len(result.matched_pairs) == 0
    assert len(result.only_in_a_unique) == 1
    assert len(result.only_in_b_unique) == 1


def test_negative_numeric_key():
    """Negative numeric keys match when equal."""
    df_a = _df([{'id': -42, 'name': 'Alice'}])
    df_b = _df([{'id': -42, 'name': 'Alice'}])
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert len(result.matched_pairs) == 1


def test_decimal_numeric_key():
    """Decimal numeric keys match when equal."""
    df_a = _df([{'id': 3.14, 'name': 'Alice'}])
    df_b = _df([{'id': 3.14, 'name': 'Alice'}])
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert len(result.matched_pairs) == 1


def test_native_date_keys_match():
    """Native date keys match when equal."""
    df_a = _df([{'d': date(2024, 1, 15), 'name': 'Alice'}])
    df_b = _df([{'d': date(2024, 1, 15), 'name': 'Alice'}])
    result = match_on_keys(df_a, df_b, keys=['d'], config=_config(keys=['d']))
    assert len(result.matched_pairs) == 1


def test_native_date_keys_different_not_match():
    """Different native date keys do not match."""
    df_a = _df([{'d': date(2024, 1, 15), 'name': 'Alice'}])
    df_b = _df([{'d': date(2024, 2, 20), 'name': 'Alice'}])
    result = match_on_keys(df_a, df_b, keys=['d'], config=_config(keys=['d']))
    assert len(result.matched_pairs) == 0
    assert len(result.only_in_a_unique) == 1
    assert len(result.only_in_b_unique) == 1


def test_native_date_vs_datetime_keys_not_match():
    """Native date vs datetime keys do not match (different string forms)."""
    df_a = _df([{'d': date(2024, 1, 15), 'name': 'Alice'}])
    df_b = _df([{'d': datetime(2024, 1, 15, 0, 0, 0), 'name': 'Alice'}])
    result = match_on_keys(df_a, df_b, keys=['d'], config=_config(keys=['d']))
    assert len(result.matched_pairs) == 0
    assert len(result.only_in_a_unique) == 1
    assert len(result.only_in_b_unique) == 1


def test_native_datetime_keys_match():
    """Native datetime keys match when equal."""
    df_a = _df([{'dt': datetime(2024, 1, 15, 9, 30, 0), 'name': 'Alice'}])
    df_b = _df([{'dt': datetime(2024, 1, 15, 9, 30, 0), 'name': 'Alice'}])
    result = match_on_keys(df_a, df_b, keys=['dt'], config=_config(keys=['dt']))
    assert len(result.matched_pairs) == 1


# ---------------------------------------------------------------------------
# Record ordering regression tests (Issue 3)
# Deliberately unsorted IDs to verify output follows source row order.
# ---------------------------------------------------------------------------


def test_matched_pairs_preserve_file_a_order():
    """Matched pairs must follow File A row order, not set iteration order."""
    df_a = _df(
        [
            {'id': 'C', 'name': 'Carol'},
            {'id': 'A', 'name': 'Alice'},
            {'id': 'B', 'name': 'Bob'},
        ]
    )
    df_b = _df(
        [
            {'id': 'A', 'name': 'Alice'},
            {'id': 'B', 'name': 'Bob'},
            {'id': 'C', 'name': 'Carol'},
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert len(result.matched_pairs) == 3
    assert result.matched_pairs[0][0]['id'] == 'C'
    assert result.matched_pairs[1][0]['id'] == 'A'
    assert result.matched_pairs[2][0]['id'] == 'B'


def test_only_in_a_preserves_file_a_order():
    """Only-in-A records must follow File A row order."""
    df_a = _df(
        [
            {'id': 'Z', 'name': 'Zara'},
            {'id': 'A', 'name': 'Alice'},
            {'id': 'M', 'name': 'Mallory'},
        ]
    )
    df_b = _df(
        [
            {'id': 'X', 'name': 'Xander'},
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert len(result.only_in_a_unique) == 3
    assert result.only_in_a_unique[0]['id'] == 'Z'
    assert result.only_in_a_unique[1]['id'] == 'A'
    assert result.only_in_a_unique[2]['id'] == 'M'


def test_only_in_b_preserves_file_b_order():
    """Only-in-B records must follow File B row order."""
    df_a = _df(
        [
            {'id': 'X', 'name': 'Xander'},
        ]
    )
    df_b = _df(
        [
            {'id': 'Z', 'name': 'Zara'},
            {'id': 'A', 'name': 'Alice'},
            {'id': 'M', 'name': 'Mallory'},
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert len(result.only_in_b_unique) == 3
    assert result.only_in_b_unique[0]['id'] == 'Z'
    assert result.only_in_b_unique[1]['id'] == 'A'
    assert result.only_in_b_unique[2]['id'] == 'M'


def test_duplicate_records_preserve_file_order():
    """Duplicate source records follow File A order, then File B order,
    with groups ordered by first appearance in each workbook."""
    df_a = _df(
        [
            {'id': 'DUP1', 'name': 'First'},
            {'id': 'DUP2', 'name': 'Second'},
            {'id': 'DUP1', 'name': 'FirstAgain'},
        ]
    )
    df_b = _df(
        [
            {'id': 'DUP2', 'name': 'BSecond'},
            {'id': 'DUP1', 'name': 'BFirst'},
        ]
    )
    result = match_on_keys(
        df_a, df_b, keys=['id'], config=_config(keys=['id'], case_sensitive=True)
    )

    # DUP2 is unambiguous (once in each file) -> matched.
    # DUP1 is ambiguous (twice in A, once in B) -> duplicates only.
    assert len(result.matched_pairs) == 1
    assert result.matched_pairs[0][0]['id'] == 'DUP2'
    assert len(result.only_in_a_unique) == 0
    assert len(result.only_in_b_unique) == 0

    # File A duplicate groups: only DUP1 (ambiguous)
    assert len(result.duplicate_groups_a) == 1
    assert result.duplicate_groups_a[0].key_values == ('DUP1',)
    assert result.duplicate_groups_a[0].occurrences == 2

    # File A duplicate source records in row order
    assert len(result.duplicate_source_records_a) == 2
    assert result.duplicate_source_records_a[0].all_column_values['name'] == 'First'
    assert result.duplicate_source_records_a[1].all_column_values['name'] == 'FirstAgain'

    # File B duplicate groups: only DUP1 (ambiguous)
    assert len(result.duplicate_groups_b) == 1
    assert result.duplicate_groups_b[0].key_values == ('DUP1',)
    assert result.duplicate_groups_b[0].occurrences == 1

    # File B duplicate source records in row order
    assert len(result.duplicate_source_records_b) == 1
    assert result.duplicate_source_records_b[0].all_column_values['name'] == 'BFirst'


def test_matching_deterministic_across_runs():
    """Repeated calls produce identical ordering — no set-iteration drift."""
    df_a = _df(
        [
            {'id': 'C', 'name': 'Carol'},
            {'id': 'A', 'name': 'Alice'},
            {'id': 'B', 'name': 'Bob'},
            {'id': 'DUP', 'name': 'Dup1'},
            {'id': 'DUP', 'name': 'Dup2'},
        ]
    )
    df_b = _df(
        [
            {'id': 'A', 'name': 'Alice'},
            {'id': 'Z', 'name': 'Zara'},
            {'id': 'DUP', 'name': 'BDup'},
        ]
    )

    result1 = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    result2 = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))

    def item_id(item):
        if isinstance(item, tuple):
            return item[0]['id']
        return item['id']

    assert [item_id(p) for p in result1.matched_pairs] == [
        item_id(p) for p in result2.matched_pairs
    ]
    assert [r['id'] for r in result1.only_in_a_unique] == [
        r['id'] for r in result2.only_in_a_unique
    ]
    assert [r['id'] for r in result1.only_in_b_unique] == [
        r['id'] for r in result2.only_in_b_unique
    ]
    assert [r.all_column_values['id'] for r in result1.duplicate_source_records_a] == [
        r.all_column_values['id'] for r in result2.duplicate_source_records_a
    ]
    assert [r.all_column_values['id'] for r in result1.duplicate_source_records_b] == [
        r.all_column_values['id'] for r in result2.duplicate_source_records_b
    ]


# ---------------------------------------------------------------------------
# Missing-value handling regression tests (Issue 4)
# None, float NaN, pd.NA, and pd.NaT must all be classified as missing keys
# consistently, never matched, and never confused with literal text.
# ---------------------------------------------------------------------------


def test_pd_na_key_classified_as_missing():
    """pd.NA key value must be classified as missing, not matched."""
    df_a = _df(
        [
            {'id': pd.NA, 'name': 'Alice'},
            {'id': '2', 'name': 'Bob'},
        ]
    )
    df_b = _df(
        [
            {'id': '2', 'name': 'Bob'},
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert len(result.missing_key_records_a) == 1
    assert result.missing_key_records_a[0].key_values == (None,)
    assert pd.isna(result.missing_key_records_a[0].all_column_values['id'])
    assert result.missing_key_records_a[0].all_column_values['name'] == 'Alice'
    assert len(result.matched_pairs) == 1
    assert len(result.only_in_a_unique) == 0


def test_pd_nat_key_classified_as_missing():
    """pd.NaT key value must be classified as missing, not matched."""
    df_a = _df(
        [
            {'id': pd.NaT, 'name': 'Alice'},
            {'id': '2', 'name': 'Bob'},
        ]
    )
    df_b = _df(
        [
            {'id': '2', 'name': 'Bob'},
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert len(result.missing_key_records_a) == 1
    assert result.missing_key_records_a[0].key_values == (None,)
    assert len(result.matched_pairs) == 1
    assert len(result.only_in_a_unique) == 0


def test_float_nan_key_classified_as_missing():
    """float NaN key value must be classified as missing, not matched."""
    df_a = _df(
        [
            {'id': float('nan'), 'name': 'Alice'},
            {'id': '2', 'name': 'Bob'},
        ]
    )
    df_b = _df(
        [
            {'id': '2', 'name': 'Bob'},
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert len(result.missing_key_records_a) == 1
    assert result.missing_key_records_a[0].key_values == (None,)
    assert len(result.matched_pairs) == 1
    assert len(result.only_in_a_unique) == 0


def test_pd_na_in_both_files_not_matched():
    """pd.NA keys in both files must not match each other; both are missing."""
    df_a = _df(
        [
            {'id': pd.NA, 'name': 'Alice'},
        ]
    )
    df_b = _df(
        [
            {'id': pd.NA, 'name': 'Bob'},
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert len(result.missing_key_records_a) == 1
    assert len(result.missing_key_records_b) == 1
    assert len(result.matched_pairs) == 0
    assert len(result.only_in_a_unique) == 0
    assert len(result.only_in_b_unique) == 0


def test_pd_na_versus_ordinary_text_not_matched():
    """pd.NA key must not match an ordinary text key."""
    df_a = _df(
        [
            {'id': pd.NA, 'name': 'Alice'},
        ]
    )
    df_b = _df(
        [
            {'id': '123', 'name': 'Bob'},
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert len(result.missing_key_records_a) == 1
    assert len(result.only_in_b_unique) == 1
    assert len(result.matched_pairs) == 0


def test_nullable_int64_column_with_pd_na():
    """Nullable Int64 column containing pd.NA must classify missing rows correctly."""
    df_a = pd.DataFrame({'id': pd.array([1, pd.NA, 3], dtype='Int64'), 'name': ['A', 'B', 'C']})
    df_b = _df(
        [
            {'id': '1', 'name': 'A'},
            {'id': '3', 'name': 'C'},
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    # pd.NA row is missing, not matched
    assert len(result.missing_key_records_a) == 1
    assert result.missing_key_records_a[0].all_column_values['name'] == 'B'
    # Valid rows matched
    assert len(result.matched_pairs) == 2
    matched_ids = {p[0]['id'] for p in result.matched_pairs}
    assert matched_ids == {1, 3}


def test_missing_key_rows_not_in_only_in():
    """Missing-key rows must not appear in only_in_a_unique or only_in_b_unique."""
    df_a = _df(
        [
            {'id': '', 'name': 'Empty'},  # empty string -> missing
            {'id': pd.NA, 'name': 'NA'},  # pd.NA -> missing
            {'id': '1', 'name': 'Valid'},  # valid -> matched
        ]
    )
    df_b = _df(
        [
            {'id': '1', 'name': 'Valid'},
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert len(result.missing_key_records_a) == 2
    assert len(result.matched_pairs) == 1
    assert len(result.only_in_a_unique) == 0
    assert all(r['id'] != '' for r in result.only_in_a_unique)


def test_row_accounting_mixed_missing_types():
    """All rows accounted for with None, pd.NA, pd.NaT and empty strings."""
    df_a = _df(
        [
            {'id': None, 'name': 'NoneRow'},
            {'id': pd.NA, 'name': 'NARow'},
            {'id': pd.NaT, 'name': 'NaTRow'},
            {'id': '', 'name': 'EmptyRow'},
            {'id': '1', 'name': 'ValidRow'},
        ]
    )
    df_b = _df(
        [
            {'id': '1', 'name': 'ValidRow'},
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    # 4 missing-key records + 1 matched = 5 total File A rows
    assert len(result.missing_key_records_a) == 4
    assert len(result.matched_pairs) == 1
    assert len(result.only_in_a_unique) == 0
    a_accounted = (
        len(result.matched_pairs) + len(result.only_in_a_unique) + len(result.missing_key_records_a)
    )
    assert a_accounted == 5


def test_literal_na_text_remains_valid_key():
    """Literal text '<NA>' must remain a valid key, not confused with pd.NA."""
    df_a = _df(
        [
            {'id': '<NA>', 'name': 'LiteralNA'},
        ]
    )
    df_b = _df(
        [
            {'id': '<NA>', 'name': 'LiteralNA'},
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert len(result.matched_pairs) == 1
    assert len(result.missing_key_records_a) == 0
    assert len(result.missing_key_records_b) == 0


def test_literal_nat_text_remains_valid_key():
    """Literal text 'NaT' must remain a valid key, not confused with pd.NaT."""
    df_a = _df(
        [
            {'id': 'NaT', 'name': 'LiteralNaT'},
        ]
    )
    df_b = _df(
        [
            {'id': 'NaT', 'name': 'LiteralNaT'},
        ]
    )
    result = match_on_keys(df_a, df_b, keys=['id'], config=_config(keys=['id']))
    assert len(result.matched_pairs) == 1
    assert len(result.missing_key_records_a) == 0
    assert len(result.missing_key_records_b) == 0
