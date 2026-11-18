"""Unit tests for the comparer module."""

from __future__ import annotations

import pandas as pd

from excel_compare.comparer import compare_rows
from excel_compare.models import ColumnPair, ComparisonConfig


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


def test_exact_match_no_diffs():
    row_a = _row({'id': '1', 'amount': '100', 'name': 'Alice'})
    row_b = _row({'id': '1', 'amount': '100', 'name': 'Alice'})
    config = _config(compare_columns=['amount', 'name'])
    diffs = compare_rows(row_a, row_b, ['amount', 'name'], config)
    assert len(diffs) == 2
    assert all(d.are_equal for d in diffs)


def test_numeric_within_tolerance_equal():
    row_a = _row({'amount': '100.000000001'})
    row_b = _row({'amount': '100.0'})
    config = _config(numeric_columns=['amount'], numeric_tolerance=1e-8)
    diffs = compare_rows(row_a, row_b, ['amount'], config)
    assert len(diffs) == 1
    assert diffs[0].are_equal is True
    # Original string values preserved
    assert diffs[0].value_a == '100.000000001'
    assert diffs[0].value_b == '100.0'


def test_numeric_outside_tolerance_mismatch():
    row_a = _row({'amount': '100.1'})
    row_b = _row({'amount': '100.0'})
    config = _config(numeric_columns=['amount'], numeric_tolerance=1e-8)
    diffs = compare_rows(row_a, row_b, ['amount'], config)
    assert len(diffs) == 1
    assert diffs[0].are_equal is False


def test_numeric_tolerance_boundary_exact():
    row_a = _row({'amount': '100.0000000005'})
    row_b = _row({'amount': '100.0'})
    config = _config(numeric_columns=['amount'], numeric_tolerance=1e-9)
    diffs = compare_rows(row_a, row_b, ['amount'], config)
    assert diffs[0].are_equal is True  # diff is 5e-10 <= 1e-9


def test_numeric_tolerance_boundary_just_outside():
    row_a = _row({'amount': '100.000000002'})
    row_b = _row({'amount': '100.0'})
    config = _config(numeric_columns=['amount'], numeric_tolerance=1e-9)
    diffs = compare_rows(row_a, row_b, ['amount'], config)
    assert diffs[0].are_equal is False  # diff is 2e-9 > 1e-9


def test_leading_zero_identifier_not_numeric():
    """'00123' must not be auto-converted to 123 — it stays text."""
    row_a = _row({'code': '00123'})
    row_b = _row({'code': '00123'})
    config = _config(compare_columns=['code'])  # not in numeric_columns
    diffs = compare_rows(row_a, row_b, ['code'], config)
    assert diffs[0].are_equal is True
    assert diffs[0].value_a == '00123'
    assert diffs[0].value_b == '00123'


def test_leading_zero_identifier_vs_numeric_value_mismatch():
    """'00123' vs '123' are different as text."""
    row_a = _row({'code': '00123'})
    row_b = _row({'code': '123'})
    config = _config(compare_columns=['code'])
    diffs = compare_rows(row_a, row_b, ['code'], config)
    assert diffs[0].are_equal is False


def test_explicit_date_parsing():
    row_a = _row({'date': '2024-01-15'})
    row_b = _row({'date': '2024-01-15'})
    config = _config(date_columns=['date'])
    diffs = compare_rows(row_a, row_b, ['date'], config)
    assert diffs[0].are_equal is True
    assert diffs[0].value_a == '2024-01-15'
    assert diffs[0].value_b == '2024-01-15'


def test_date_mismatch():
    row_a = _row({'date': '2024-01-15'})
    row_b = _row({'date': '2024-02-20'})
    config = _config(date_columns=['date'])
    diffs = compare_rows(row_a, row_b, ['date'], config)
    assert diffs[0].are_equal is False


def test_ambiguous_date_format_unparseable_mismatch():
    """Ambiguous/unparseable date strings are treated as mismatch."""
    row_a = _row({'date': 'not-a-date'})
    row_b = _row({'date': '2024-01-15'})
    config = _config(date_columns=['date'])
    diffs = compare_rows(row_a, row_b, ['date'], config)
    assert diffs[0].are_equal is False


def test_unconfigured_column_uses_text_comparison():
    """'100' vs '100.0' as text are NOT equal — no auto numeric conversion."""
    row_a = _row({'amount': '100'})
    row_b = _row({'amount': '100.0'})
    config = _config(compare_columns=['amount'])  # not in numeric_columns
    diffs = compare_rows(row_a, row_b, ['amount'], config)
    assert diffs[0].are_equal is False


def test_missing_vs_missing_equal():
    row_a = _row({'amount': None})
    row_b = _row({'amount': None})
    config = _config(compare_columns=['amount'])
    diffs = compare_rows(row_a, row_b, ['amount'], config)
    assert diffs[0].are_equal is True


def test_missing_vs_present_mismatch():
    row_a = _row({'amount': None})
    row_b = _row({'amount': '100'})
    config = _config(compare_columns=['amount'])
    diffs = compare_rows(row_a, row_b, ['amount'], config)
    assert diffs[0].are_equal is False


def test_present_vs_missing_mismatch():
    row_a = _row({'amount': '100'})
    row_b = _row({'amount': None})
    config = _config(compare_columns=['amount'])
    diffs = compare_rows(row_a, row_b, ['amount'], config)
    assert diffs[0].are_equal is False


def test_empty_vs_missing_not_equal_by_default():
    """Empty string '' vs missing None are NOT equal by default."""
    row_a = _row({'amount': ''})
    row_b = _row({'amount': None})
    config = _config(compare_columns=['amount'])
    diffs = compare_rows(row_a, row_b, ['amount'], config)
    assert diffs[0].are_equal is False


def test_empty_vs_empty_equal():
    row_a = _row({'amount': ''})
    row_b = _row({'amount': ''})
    config = _config(compare_columns=['amount'])
    diffs = compare_rows(row_a, row_b, ['amount'], config)
    assert diffs[0].are_equal is True


def test_case_insensitive_text_match():
    row_a = _row({'name': 'ALICE'})
    row_b = _row({'name': 'alice'})
    config = _config(compare_columns=['name'], case_sensitive=False)
    diffs = compare_rows(row_a, row_b, ['name'], config)
    assert diffs[0].are_equal is True


def test_case_sensitive_text_mismatch():
    row_a = _row({'name': 'ALICE'})
    row_b = _row({'name': 'alice'})
    config = _config(compare_columns=['name'], case_sensitive=True)
    diffs = compare_rows(row_a, row_b, ['name'], config)
    assert diffs[0].are_equal is False


def test_whitespace_normalization_text_match():
    row_a = _row({'name': ' Alice '})
    row_b = _row({'name': 'Alice'})
    config = _config(
        compare_columns=['name'],
        normalize_whitespace=True,
        case_sensitive=True,
    )
    diffs = compare_rows(row_a, row_b, ['name'], config)
    assert diffs[0].are_equal is True


def test_original_values_preserved_in_celldiff():
    """CellDiff stores original values, not normalized/converted ones."""
    row_a = _row({'name': ' Alice ', 'amount': '100.5'})
    row_b = _row({'name': 'alice', 'amount': '100.50'})
    config = _config(
        compare_columns=['name', 'amount'],
        case_sensitive=False,
        normalize_whitespace=True,
    )
    diffs = compare_rows(row_a, row_b, ['name', 'amount'], config)
    name_diff = diffs[0]
    amount_diff = diffs[1]
    assert name_diff.value_a == ' Alice '
    assert name_diff.value_b == 'alice'
    assert amount_diff.value_a == '100.5'
    assert amount_diff.value_b == '100.50'


def test_row_series_not_modified():
    row_a = _row({'name': ' Alice ', 'amount': '100'})
    row_b = _row({'name': 'Alice', 'amount': '100'})
    original_a = row_a.copy()
    original_b = row_b.copy()
    config = _config(
        compare_columns=['name', 'amount'],
        normalize_whitespace=True,
        case_sensitive=False,
    )
    compare_rows(row_a, row_b, ['name', 'amount'], config)
    pd.testing.assert_series_equal(row_a, original_a)
    pd.testing.assert_series_equal(row_b, original_b)


def test_renamed_compare_pair_matching():
    """Compare pair with different column names on each side."""
    row_a = _row({'Make': 'Toyota', 'Model': 'Camry'})
    row_b = _row({'Camera Make': 'Toyota', 'Model': 'Camry'})
    config = _config(
        compare_pairs=[ColumnPair('Make', 'Camera Make'), ColumnPair('Model', 'Model')]
    )
    diffs = compare_rows(row_a, row_b, config.compare_pairs, config)
    assert len(diffs) == 2
    assert all(d.are_equal for d in diffs)
    assert diffs[0].column_a == 'Make'
    assert diffs[0].column_b == 'Camera Make'


def test_renamed_compare_pair_mismatch():
    """Renamed compare pair with different values is a mismatch."""
    row_a = _row({'Make': 'Toyota'})
    row_b = _row({'Camera Make': 'Honda'})
    config = _config(compare_pairs=[ColumnPair('Make', 'Camera Make')])
    diffs = compare_rows(row_a, row_b, config.compare_pairs, config)
    assert len(diffs) == 1
    assert diffs[0].are_equal is False
    assert diffs[0].column_a == 'Make'
    assert diffs[0].column_b == 'Camera Make'
    assert diffs[0].value_a == 'Toyota'
    assert diffs[0].value_b == 'Honda'


def test_numeric_pair_with_different_names():
    """Numeric comparison on explicitly configured renamed pair."""
    row_a = _row({'Amount': '100.5'})
    row_b = _row({'Cost': '100.500000001'})
    config = _config(
        compare_pairs=[ColumnPair('Amount', 'Cost')],
        numeric_pairs=[ColumnPair('Amount', 'Cost')],
        numeric_tolerance=1e-8,
    )
    diffs = compare_rows(row_a, row_b, config.compare_pairs, config)
    assert len(diffs) == 1
    assert diffs[0].are_equal is True


def test_date_pair_with_different_names():
    """Date comparison on explicitly configured renamed pair."""
    row_a = _row({'Date A': '2024-01-15'})
    row_b = _row({'Date B': '2024-01-15'})
    config = _config(
        compare_pairs=[ColumnPair('Date A', 'Date B')],
        date_pairs=[ColumnPair('Date A', 'Date B')],
    )
    diffs = compare_rows(row_a, row_b, config.compare_pairs, config)
    assert len(diffs) == 1
    assert diffs[0].are_equal is True


def test_column_pair_not_in_compare_pairs_uses_text():
    """Column pair not in numeric_pairs uses text comparison, not numeric."""
    row_a = _row({'Amount': '100'})
    row_b = _row({'Cost': '100.0'})
    config = _config(
        compare_pairs=[ColumnPair('Amount', 'Cost')],
        # Amount/Cost NOT in numeric_pairs
    )
    diffs = compare_rows(row_a, row_b, config.compare_pairs, config)
    assert len(diffs) == 1
    assert diffs[0].are_equal is False  # '100' != '100.0' as text


def test_empty_compare_pairs_no_fields_compared():
    """Empty compare_pairs results in no cell diffs (not treated as equal)."""
    row_a = _row({'id': '1', 'name': 'Alice'})
    row_b = _row({'id': '1', 'name': 'Bob'})
    config = _config(compare_pairs=[])
    diffs = compare_rows(row_a, row_b, [], config)
    assert len(diffs) == 0
