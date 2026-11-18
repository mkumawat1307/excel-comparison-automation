"""Unit tests for the normalizer module."""

from __future__ import annotations

import pandas as pd

from excel_compare.normalizer import canonical_missing, normalize_dataframe


def test_canonical_missing_none() -> None:
    assert canonical_missing(None) is True


def test_canonical_missing_nan() -> None:
    assert canonical_missing(float('nan')) is True


def test_canonical_missing_pd_na() -> None:
    assert canonical_missing(pd.NA) is True


def test_canonical_missing_pd_nat() -> None:
    assert canonical_missing(pd.NaT) is True


def test_canonical_missing_empty_string_is_not_missing() -> None:
    assert canonical_missing('') is False


def test_canonical_missing_regular_value() -> None:
    assert canonical_missing('Alice') is False
    assert canonical_missing(0) is False
    assert canonical_missing(100) is False


def test_canonical_missing_non_scalar_safe() -> None:
    # Non-scalar values must not raise ambiguous truth-value errors.
    assert canonical_missing([1, 2, 3]) is False
    assert canonical_missing(pd.Series([1, None])) is False
    assert canonical_missing(pd.DataFrame({'a': [1]})) is False


def test_whitespace_normalization() -> None:
    df = pd.DataFrame({'name': [' Alice ', 'Bob  ', '\tCharlie\n']})
    result = normalize_dataframe(df, normalize_whitespace=True, case_sensitive=True)
    assert result.iloc[0]['name'] == 'Alice'
    assert result.iloc[1]['name'] == 'Bob'
    assert result.iloc[2]['name'] == 'Charlie'


def test_whitespace_not_normalized_when_disabled() -> None:
    df = pd.DataFrame({'name': [' Alice ']})
    result = normalize_dataframe(df, normalize_whitespace=False, case_sensitive=True)
    assert result.iloc[0]['name'] == ' Alice '


def test_case_normalization_default() -> None:
    df = pd.DataFrame({'name': ['ALICE', 'Bob']})
    result = normalize_dataframe(df, case_sensitive=False)
    assert result.iloc[0]['name'] == 'alice'
    assert result.iloc[1]['name'] == 'bob'


def test_case_sensitive_preserves_case() -> None:
    df = pd.DataFrame({'name': ['ALICE', 'Bob']})
    result = normalize_dataframe(df, case_sensitive=True)
    assert result.iloc[0]['name'] == 'ALICE'
    assert result.iloc[1]['name'] == 'Bob'


def test_missing_values_pass_through() -> None:
    df = pd.DataFrame({'name': ['Alice', None, float('nan'), pd.NA]})
    result = normalize_dataframe(df)
    assert canonical_missing(result.iloc[1]['name'])
    assert canonical_missing(result.iloc[2]['name'])
    assert canonical_missing(result.iloc[3]['name'])


def test_empty_string_preserved_by_default() -> None:
    df = pd.DataFrame({'name': ['Alice', '']})
    result = normalize_dataframe(df)
    assert result.iloc[0]['name'] == 'alice'
    assert result.iloc[1]['name'] == ''


def test_treat_empty_as_missing_converts() -> None:
    df = pd.DataFrame({'name': ['Alice', '']})
    result = normalize_dataframe(df, treat_empty_as_missing=True)
    assert canonical_missing(result.iloc[1]['name'])


def test_numeric_values_preserved() -> None:
    df = pd.DataFrame({'amount': [100, 200.5, None], 'id': [1, 2, 3]})
    result = normalize_dataframe(df)
    assert result.iloc[0]['amount'] == 100
    assert result.iloc[1]['amount'] == 200.5
    # Numeric dtype unchanged
    assert pd.api.types.is_numeric_dtype(result['amount'])
    assert pd.api.types.is_numeric_dtype(result['id'])


def test_text_identifiers_keep_leading_zeros() -> None:
    df = pd.DataFrame({'code': ['00123', '00456']})
    result = normalize_dataframe(df)
    assert result.iloc[0]['code'] == '00123'
    assert result.iloc[1]['code'] == '00456'
    assert pd.api.types.is_string_dtype(result['code'])


def test_date_values_preserved() -> None:
    dates = pd.to_datetime(['2024-01-15', '2024-02-20'])
    df = pd.DataFrame({'date': dates})
    result = normalize_dataframe(df)
    assert pd.api.types.is_datetime64_any_dtype(result['date'])
    assert result.iloc[0]['date'] == dates[0]


def test_original_dataframe_unchanged() -> None:
    df = pd.DataFrame({'name': [' Alice ', 'ALICE'], 'amount': [100, 200]})
    original = df.copy()
    normalize_dataframe(df, normalize_whitespace=True, case_sensitive=False)
    pd.testing.assert_frame_equal(df, original)


def test_normalize_returns_copy_not_same_object() -> None:
    df = pd.DataFrame({'name': ['Alice']})
    result = normalize_dataframe(df)
    assert result is not df
