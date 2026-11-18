"""Per-cell comparison.

Compares matched rows cell-by-cell with numeric tolerance and
missing-value handling.

Type resolution rules:
- Numeric comparison is performed ONLY on column pairs explicitly listed in
  `config.numeric_pairs`. String values like "00123" are never
  auto-converted to numbers.
- Date comparison is performed ONLY on column pairs explicitly listed in
  `config.date_pairs`, using pd.to_datetime with dayfirst=False.
- All other compared pairs use text comparison.
- Original string values are preserved in CellDiff for reporting.
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Union

import pandas as pd

from .models import CellDiff, ColumnPair, ComparisonConfig


def _as_column_pairs(
    column_pairs: List[Union[str, ColumnPair]],
) -> List[ColumnPair]:
    """Normalize column specifications to ColumnPair objects.

    Accepts bare strings (legacy: same-name pairs) or ColumnPair objects.
    """
    result: List[ColumnPair] = []
    for item in column_pairs:
        if isinstance(item, ColumnPair):
            result.append(item)
        else:
            result.append(ColumnPair(file_a_column=item, file_b_column=item))
    return result


def _is_missing(value: Any) -> bool:
    """Return True if a scalar value is missing (None, NaN, pd.NA, pd.NaT).

    Empty strings are NOT treated as missing by default.
    """
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    if isinstance(value, type(pd.NA)) or value is pd.NA:
        return True
    if isinstance(value, pd.Timestamp) and pd.isna(value):
        return True
    return False


def _empty_string(value: Any) -> bool:
    """Return True if value is an empty string."""
    return isinstance(value, str) and value.strip() == ''


def _normalize_text(
    value: Any,
    case_sensitive: bool,
    normalize_whitespace: bool,
) -> str:
    """Normalize a text value for comparison without mutating the original."""
    if not isinstance(value, str):
        return str(value)
    result = value
    if normalize_whitespace:
        result = result.strip()
    if not case_sensitive:
        result = result.lower()
    return result


def _try_parse_date(value: Any) -> Optional[pd.Timestamp]:
    """Parse a string as a date using documented rules (dayfirst=False).

    Returns None if parsing fails.
    """
    if not isinstance(value, str):
        return None
    stripped = value.strip()
    if stripped == '':
        return None
    try:
        return pd.to_datetime(stripped, dayfirst=False)
    except (ValueError, TypeError):
        return None


def _try_parse_numeric(value: Any) -> Optional[float]:
    """Parse a string as a float.

    Returns None if parsing fails.
    """
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        stripped = value.strip()
        if stripped == '':
            return None
        try:
            return float(stripped)
        except ValueError:
            return None
    return None


def compare_rows(
    row_a: pd.Series,
    row_b: pd.Series,
    column_pairs: List[ColumnPair],
    config: ComparisonConfig,
    sheet: str = '',
    row_key: str = '',
    excel_row_mapping_a: Optional[Dict[int, int]] = None,
    excel_row_mapping_b: Optional[Dict[int, int]] = None,
) -> List[CellDiff]:
    """Compare two rows cell-by-cell across the given column pairs.

    Original values are preserved in every CellDiff. Neither row is
    modified.

    Parameters
    ----------
    row_a : pd.Series
        Row from File A.
    row_b : pd.Series
        Row from File B.
    column_pairs : list[ColumnPair]
        Column pairs to compare. Each pair specifies File A column and
        File B column names.
    config : ComparisonConfig
        Comparison settings (numeric_pairs, date_pairs, tolerance,
        case sensitivity, whitespace normalization).
    sheet : str
        Sheet name for the CellDiff record.
    row_key : str
        Composite key string for the CellDiff record.

    Returns
    -------
    list[CellDiff]
        One CellDiff per compared column pair. are_equal=True means the
        values match under the configured rules.
    """
    numeric_pair_set = {(p.file_a_column, p.file_b_column) for p in config.numeric_pairs}
    if not numeric_pair_set and config.numeric_columns:
        numeric_pair_set = {(c, c) for c in config.numeric_columns}
    date_pair_set = {(p.file_a_column, p.file_b_column) for p in config.date_pairs}
    if not date_pair_set and config.date_columns:
        date_pair_set = {(c, c) for c in config.date_columns}
    case_sensitive = config.case_sensitive
    normalize_ws = config.normalize_whitespace
    tolerance = config.numeric_tolerance

    diffs: List[CellDiff] = []

    resolved_pairs = _as_column_pairs(column_pairs)

    # Resolve physical Excel row numbers from DataFrame index for display.
    excel_row_a = None
    excel_row_b = None
    if excel_row_mapping_a and hasattr(row_a, 'name') and row_a.name is not None:
        excel_row_a = excel_row_mapping_a.get(int(row_a.name))
    if excel_row_mapping_b and hasattr(row_b, 'name') and row_b.name is not None:
        excel_row_b = excel_row_mapping_b.get(int(row_b.name))

    for pair in resolved_pairs:
        col_a = pair.file_a_column
        col_b = pair.file_b_column
        v_a = row_a.get(col_a)
        v_b = row_b.get(col_b)

        missing_a = _is_missing(v_a)
        missing_b = _is_missing(v_b)
        empty_a = _empty_string(v_a)
        empty_b = _empty_string(v_b)

        # Missing vs missing -> equal
        if missing_a and missing_b:
            diffs.append(
                CellDiff(
                    column_a=col_a,
                    column_b=col_b,
                    row_key=row_key,
                    sheet=sheet,
                    value_a=v_a,
                    value_b=v_b,
                    are_equal=True,
                    excel_row_a=excel_row_a,
                    excel_row_b=excel_row_b,
                )
            )
            continue

        # One missing, one present -> mismatch (empty string vs missing
        # is also a mismatch unless treat_empty_as_missing is configured;
        # that flag lives in the normalizer, so here we treat '' as a
        # real value, not missing).
        if missing_a or missing_b:
            diffs.append(
                CellDiff(
                    column_a=col_a,
                    column_b=col_b,
                    row_key=row_key,
                    sheet=sheet,
                    value_a=v_a,
                    value_b=v_b,
                    are_equal=False,
                    excel_row_a=excel_row_a,
                    excel_row_b=excel_row_b,
                )
            )
            continue

        # Both empty strings -> equal (text comparison)
        if empty_a and empty_b:
            diffs.append(
                CellDiff(
                    column_a=col_a,
                    column_b=col_b,
                    row_key=row_key,
                    sheet=sheet,
                    value_a=v_a,
                    value_b=v_b,
                    are_equal=True,
                    excel_row_a=excel_row_a,
                    excel_row_b=excel_row_b,
                )
            )
            continue

        # Numeric comparison: only on explicitly configured pairs.
        if (col_a, col_b) in numeric_pair_set:
            n_a = _try_parse_numeric(v_a)
            n_b = _try_parse_numeric(v_b)
            if n_a is None or n_b is None:
                # One side couldn't be parsed as numeric -> mismatch,
                # report original values.
                diffs.append(
                    CellDiff(
                        column_a=col_a,
                        column_b=col_b,
                        row_key=row_key,
                        sheet=sheet,
                        value_a=v_a,
                        value_b=v_b,
                        are_equal=False,
                        excel_row_a=excel_row_a,
                        excel_row_b=excel_row_b,
                    )
                )
                continue
            are_equal = abs(n_a - n_b) <= tolerance
            diffs.append(
                CellDiff(
                    column_a=col_a,
                    column_b=col_b,
                    row_key=row_key,
                    sheet=sheet,
                    value_a=v_a,
                    value_b=v_b,
                    are_equal=are_equal,
                    excel_row_a=excel_row_a,
                    excel_row_b=excel_row_b,
                )
            )
            continue

        # Date comparison: only on explicitly configured pairs.
        if (col_a, col_b) in date_pair_set:
            d_a = _try_parse_date(v_a)
            d_b = _try_parse_date(v_b)
            if d_a is None or d_b is None:
                diffs.append(
                    CellDiff(
                        column_a=col_a,
                        column_b=col_b,
                        row_key=row_key,
                        sheet=sheet,
                        value_a=v_a,
                        value_b=v_b,
                        are_equal=False,
                        excel_row_a=excel_row_a,
                        excel_row_b=excel_row_b,
                    )
                )
                continue
            are_equal = d_a == d_b
            diffs.append(
                CellDiff(
                    column_a=col_a,
                    column_b=col_b,
                    row_key=row_key,
                    sheet=sheet,
                    value_a=v_a,
                    value_b=v_b,
                    are_equal=are_equal,
                    excel_row_a=excel_row_a,
                    excel_row_b=excel_row_b,
                )
            )
            continue

        # Text comparison (default for all unconfigured pairs).
        norm_a = _normalize_text(v_a, case_sensitive, normalize_ws)
        norm_b = _normalize_text(v_b, case_sensitive, normalize_ws)
        are_equal = norm_a == norm_b
        diffs.append(
            CellDiff(
                column_a=col_a,
                column_b=col_b,
                row_key=row_key,
                sheet=sheet,
                value_a=v_a,
                value_b=v_b,
                are_equal=are_equal,
                excel_row_a=excel_row_a,
                excel_row_b=excel_row_b,
            )
        )

    return diffs
