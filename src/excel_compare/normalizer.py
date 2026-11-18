"""Value normalization.

Handles whitespace stripping, case handling, and missing-value detection
before comparison. Non-mutating: the original DataFrame is never changed.
Original data types (numbers, dates, text identifiers with leading zeros)
are preserved.
"""

from __future__ import annotations

from typing import Any

import pandas as pd


def canonical_missing(value: Any) -> bool:
    """Return True if a scalar value should be treated as missing.

    Handles None, NaN, pandas.NA, and pandas.NaT consistently.
    Empty strings are NOT treated as missing by default, preserving the
    distinction between empty cells and missing values.

    Non-scalar values (arrays, Series, DataFrames) return False to avoid
    ambiguous truth-value errors from pd.isna().
    """
    if not pd.api.types.is_scalar(value):
        return False
    return bool(pd.isna(value))


def normalize_dataframe(
    df: pd.DataFrame,
    case_sensitive: bool = False,
    normalize_whitespace: bool = True,
    treat_empty_as_missing: bool = False,
) -> pd.DataFrame:
    """Return a normalized copy of the DataFrame.

    Only object/string columns are transformed. Numeric and datetime
    columns are left untouched so original types are preserved.

    Parameters
    ----------
    df : pd.DataFrame
        Source DataFrame (not modified).
    case_sensitive : bool
        If False (default), text is lowercased for comparison.
    normalize_whitespace : bool
        If True (default), leading/trailing whitespace is stripped.
    treat_empty_as_missing : bool
        If True, empty strings are converted to NaN. Default False keeps
        empty strings distinct from missing values.

    Returns
    -------
    pd.DataFrame
        A new DataFrame with normalization applied. The input is unchanged.
    """
    result = df.copy()

    if treat_empty_as_missing:
        result = result.replace('', pd.NA)

    for column in result.columns:
        if not pd.api.types.is_string_dtype(result[column]):
            continue

        series = result[column]
        if normalize_whitespace:
            series = series.str.strip()
        if not case_sensitive:
            series = series.str.lower()
        result[column] = series

    return result
