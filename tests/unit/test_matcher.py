"""Unit tests for the matcher module (Phase 4)."""

from __future__ import annotations

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
