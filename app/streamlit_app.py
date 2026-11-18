"""Streamlit entry point for Excel Comparison Automation.

Non-technical business users upload two Excel files, configure matching keys
and comparison columns, run the comparison, and download a detailed XLSX report.
"""

from __future__ import annotations

import hashlib
import io
import logging
import os
import tempfile
import uuid
from typing import Optional

import streamlit as st

from excel_compare.comparer import compare_rows
from excel_compare.config import (
    DEFAULT_CASE_SENSITIVE,
    DEFAULT_MAX_FILE_SIZE_MB,
    DEFAULT_NORMALIZE_WHITESPACE,
    DEFAULT_NUMERIC_TOLERANCE,
)
from excel_compare.exceptions import (
    DuplicateKeyError,
    InvalidKeySelectionError,
    MappingValidationError,
    SecurityError,
)
from excel_compare.loader import list_sheets, load_sheet_with_metadata
from excel_compare.matcher import discover_auto_pairs, match_on_keys
from excel_compare.models import ColumnPair, ComparisonConfig
from excel_compare.reporter import (
    build_report,
    file_a_column_header,
    file_b_column_header,
)
from excel_compare.security import validate_uploaded_file
from excel_compare.writer import write_report

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Session-state initialization
# ---------------------------------------------------------------------------


def _init_session_state() -> None:
    """Initialize all session-state keys."""
    defaults = {
        'file_a_bytes': None,
        'file_b_bytes': None,
        'file_a_name': None,
        'file_b_name': None,
        'file_a_hash': None,
        'file_b_hash': None,
        'file_a_obj': None,
        'file_b_obj': None,
        'sheet_a': None,
        'sheet_b': None,
        'sheets_a': [],
        'sheets_b': [],
        'loaded_a': None,
        'loaded_b': None,
        'columns_a': None,
        'columns_b': None,
        'key_pair_entries': [],
        'compare_pair_entries': [],
        'numeric_pairs': [],
        'date_pairs': [],
        'comparison_result': None,
        'report_bytes': None,
        'current_step': 1,
        '_last_key_pairs': None,
        '_last_compare_pairs': None,
        '_last_numeric_pairs': None,
        '_last_date_pairs': None,
        '_auto_added_key_ids': set(),
        '_auto_added_compare_ids': set(),
        '_session_nonce': 0,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def _reset_from_upload() -> None:
    """Reset all downstream state when either uploaded file changes."""
    st.session_state['sheet_a'] = None
    st.session_state['sheet_b'] = None
    st.session_state['sheets_a'] = []
    st.session_state['sheets_b'] = []
    st.session_state['loaded_a'] = None
    st.session_state['loaded_b'] = None
    st.session_state['columns_a'] = None
    st.session_state['columns_b'] = None
    st.session_state['key_pair_entries'] = []
    st.session_state['compare_pair_entries'] = []
    st.session_state['numeric_pairs'] = []
    st.session_state['date_pairs'] = []
    st.session_state['_auto_added_key_ids'] = set()
    st.session_state['_auto_added_compare_ids'] = set()
    st.session_state['comparison_result'] = None
    st.session_state['report_bytes'] = None


def _reset_from_worksheet() -> None:
    """Reset all downstream state when either worksheet changes."""
    st.session_state['loaded_a'] = None
    st.session_state['loaded_b'] = None
    st.session_state['columns_a'] = None
    st.session_state['columns_b'] = None
    st.session_state['key_pair_entries'] = []
    st.session_state['compare_pair_entries'] = []
    st.session_state['numeric_pairs'] = []
    st.session_state['date_pairs'] = []
    st.session_state['_auto_added_key_ids'] = set()
    st.session_state['_auto_added_compare_ids'] = set()
    st.session_state['comparison_result'] = None
    st.session_state['report_bytes'] = None


def _reset_from_mapping() -> None:
    """Reset all downstream state when key mappings change."""
    st.session_state['compare_pair_entries'] = []
    st.session_state['numeric_pairs'] = []
    st.session_state['date_pairs'] = []
    st.session_state['comparison_result'] = None
    st.session_state['report_bytes'] = None


def _reset_from_comparison() -> None:
    """Reset results when comparison mappings/types change."""
    st.session_state['comparison_result'] = None
    st.session_state['report_bytes'] = None


def _clear_session_data() -> None:
    """Clear all application-owned session data and return to Step 1.

    Intended to run as a Streamlit ``on_click`` callback, which executes before
    widgets are instantiated on the following run. This drops the references
    the app holds to uploaded bytes, uploaded file objects, loaded worksheets,
    DataFrames, mappings, comparison results, report bytes, and widget values.

    The uploader nonce is preserved and incremented so the file-uploader widget
    keys rotate; previously uploaded files therefore cannot reappear on the
    next rerun.

    Note: this releases Python references. It is not a guarantee of secure
    physical memory erasure.
    """
    nonce = st.session_state.get('_session_nonce', 0)
    st.session_state.clear()
    st.session_state['_session_nonce'] = nonce + 1


def _request_clear(confirm_key: str) -> None:
    """``on_click`` callback: clear data only after explicit confirmation."""
    if st.session_state.get(confirm_key):
        _clear_session_data()
    else:
        st.session_state['_clear_blocked'] = True


def _compute_hash(data: bytes) -> str:
    """Return a SHA-256 hex digest of file bytes."""
    return hashlib.sha256(data).hexdigest()


def _validate_upload(uploaded_file) -> Optional[bytes]:
    """Validate an uploaded file and return its bytes.

    Returns None if validation fails (error already surfaced via st.error).
    """
    if uploaded_file is None:
        return None

    raw = uploaded_file.read()

    # Size check
    if len(raw) > DEFAULT_MAX_FILE_SIZE_MB * 1024 * 1024:
        st.error(
            f'File "{uploaded_file.name}" exceeds maximum size of {DEFAULT_MAX_FILE_SIZE_MB} MB.'
        )
        return None

    # Extension check
    name = uploaded_file.name
    if not name.lower().endswith('.xlsx'):
        st.error(f'File "{name}" is not an .xlsx file.')
        return None

    # Readability check via security.validate_uploaded_file
    # Write bytes to a temp file so validate_uploaded_file can open it by path.
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            tmp.write(raw)
            tmp_path = tmp.name
        validate_uploaded_file(tmp_path)
    except SecurityError as exc:
        st.error(str(exc))
        return None
    finally:
        if tmp_path is not None:
            try:
                os.unlink(tmp_path)
            except OSError as exc:
                logger.warning('Failed to delete temporary upload file %s: %s', tmp_path, exc)

    return raw


def _load_worksheet(uploaded_file, sheet_name: str):
    """Load a worksheet from an uploaded file's bytes."""
    raw = (
        st.session_state.get('file_a_bytes')
        if uploaded_file is st.session_state.get('file_a_obj')
        else st.session_state.get('file_b_bytes')
    )
    # Fall back to reading from the uploaded file object directly
    if raw is None:
        raw = uploaded_file.read()
    buffer = io.BytesIO(raw)
    return load_sheet_with_metadata(buffer, sheet_name, track_excel_row=True)


def _build_comparison_config() -> ComparisonConfig:
    """Build a ComparisonConfig from session-state mappings."""
    key_pair_entries = st.session_state.get('key_pair_entries', [])
    compare_pair_entries = st.session_state.get('compare_pair_entries', [])
    numeric_pairs = st.session_state.get('numeric_pairs', [])
    date_pairs = st.session_state.get('date_pairs', [])
    key_pairs = [
        ColumnPair(file_a_column=e['file_a'], file_b_column=e['file_b'])
        for e in key_pair_entries
        if e['file_a'] is not None and e['file_b'] is not None
    ]
    compare_pairs = [
        ColumnPair(file_a_column=e['file_a'], file_b_column=e['file_b'])
        for e in compare_pair_entries
        if e['file_a'] is not None and e['file_b'] is not None
    ]
    numeric_pair_list = [ColumnPair(file_a_column=a, file_b_column=b) for a, b in numeric_pairs]
    date_pair_list = [ColumnPair(file_a_column=a, file_b_column=b) for a, b in date_pairs]
    return ComparisonConfig(
        keys=[],
        compare_columns=[],
        numeric_columns=[],
        date_columns=[],
        key_pairs=key_pairs,
        compare_pairs=compare_pairs,
        numeric_pairs=numeric_pair_list,
        date_pairs=date_pair_list,
        case_sensitive=DEFAULT_CASE_SENSITIVE,
        normalize_whitespace=DEFAULT_NORMALIZE_WHITESPACE,
        numeric_tolerance=DEFAULT_NUMERIC_TOLERANCE,
        duplicate_policy='error',
        max_file_size_mb=DEFAULT_MAX_FILE_SIZE_MB,
    )


def _run_comparison() -> None:
    """Run the full comparison pipeline and store results in session state."""
    loaded_a = st.session_state.get('loaded_a')
    loaded_b = st.session_state.get('loaded_b')
    key_pair_entries = st.session_state.get('key_pair_entries', [])
    compare_pair_entries = st.session_state.get('compare_pair_entries', [])

    if loaded_a is None or loaded_b is None:
        st.error('Please load both worksheets before running the comparison.')
        return

    # Validate key pairs: reject partial entries (one column selected, other empty).
    valid_key_entries = [
        e for e in key_pair_entries if e['file_a'] is not None and e['file_b'] is not None
    ]
    partial_key_entries = [
        e for e in key_pair_entries if (e['file_a'] is None) != (e['file_b'] is None)
    ]
    if partial_key_entries:
        st.error(
            'Incomplete key mapping(s): both File A and File B columns must be selected '
            'for every key pair. Remove or complete the incomplete pair(s) before running.'
        )
        return
    if not valid_key_entries:
        st.error('At least one valid matching key pair is required.')
        return

    config = _build_comparison_config()

    try:
        match_result = match_on_keys(
            df_a=loaded_a.df,
            df_b=loaded_b.df,
            key_pairs=[ColumnPair(e['file_a'], e['file_b']) for e in valid_key_entries],
            config=config,
            excel_row_mapping_a=loaded_a.excel_row_mapping,
            excel_row_mapping_b=loaded_b.excel_row_mapping,
        )
    except (InvalidKeySelectionError, MappingValidationError, DuplicateKeyError) as exc:
        st.error(str(exc))
        return

    # Compare each matched pair
    cell_diffs_by_pair = []
    excel_row_mapping_a = loaded_a.excel_row_mapping
    excel_row_mapping_b = loaded_b.excel_row_mapping

    # Validate comparison pairs: reject partial entries.
    valid_compare_entries = [
        e for e in compare_pair_entries if e['file_a'] is not None and e['file_b'] is not None
    ]
    partial_compare_entries = [
        e for e in compare_pair_entries if (e['file_a'] is None) != (e['file_b'] is None)
    ]
    if partial_compare_entries:
        st.error(
            'Incomplete comparison mapping(s): both File A and File B columns must be '
            'selected for every comparison pair. Remove or complete the incomplete '
            'pair(s) before running.'
        )
        return

    for row_a, row_b in match_result.matched_pairs:
        diffs = compare_rows(
            row_a=row_a,
            row_b=row_b,
            column_pairs=[ColumnPair(e['file_a'], e['file_b']) for e in valid_compare_entries],
            config=config,
            sheet=loaded_a.sheet_name,
            row_key='',
            excel_row_mapping_a=excel_row_mapping_a,
            excel_row_mapping_b=excel_row_mapping_b,
        )
        cell_diffs_by_pair.append(diffs)

    comparison_result = build_report(
        match_result=match_result,
        cell_diffs_by_pair=cell_diffs_by_pair,
        config=config,
        excel_row_mapping_a=excel_row_mapping_a,
        excel_row_mapping_b=excel_row_mapping_b,
    )

    # Write report to a temp file, then read bytes
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            write_report(
                result=comparison_result,
                output_path=tmp.name,
                config=config,
                df_a=loaded_a.df,
                df_b=loaded_b.df,
            )
            tmp_path = tmp.name
        with open(tmp_path, 'rb') as f:
            report_bytes = f.read()
    finally:
        if tmp_path is not None:
            try:
                os.unlink(tmp_path)
            except OSError as exc:
                logger.warning('Failed to delete temporary report file %s: %s', tmp_path, exc)

    st.session_state['comparison_result'] = comparison_result
    st.session_state['report_bytes'] = report_bytes


def _render_step1_upload() -> None:
    """Step 1: Upload files."""
    st.subheader('Step 1: Upload Excel Files')
    st.caption(
        'Upload two Excel (.xlsx) files to compare. Maximum file size: 50 MB per .xlsx file.'
    )

    col_a, col_b = st.columns(2)

    nonce = st.session_state.get('_session_nonce', 0)

    with col_a:
        file_a = st.file_uploader(
            'File A (old/existing data)',
            type=['xlsx'],
            key=f'uploader_file_a_{nonce}',
        )
        if file_a is not None and st.session_state.get('file_a_obj') is not file_a:
            raw = _validate_upload(file_a)
            if raw is not None:
                st.session_state['file_a_obj'] = file_a
                st.session_state['file_a_bytes'] = raw
                st.session_state['file_a_name'] = file_a.name
                st.session_state['file_a_hash'] = _compute_hash(raw)
                _reset_from_upload()

    with col_b:
        file_b = st.file_uploader(
            'File B (new/current data)',
            type=['xlsx'],
            key=f'uploader_file_b_{nonce}',
        )
        if file_b is not None and st.session_state.get('file_b_obj') is not file_b:
            raw = _validate_upload(file_b)
            if raw is not None:
                st.session_state['file_b_obj'] = file_b
                st.session_state['file_b_bytes'] = raw
                st.session_state['file_b_name'] = file_b.name
                st.session_state['file_b_hash'] = _compute_hash(raw)
                _reset_from_upload()

    file_a = st.session_state.get('file_a_obj')
    file_b = st.session_state.get('file_b_obj')

    if file_a is not None and file_b is not None:
        st.success(
            f'Both files uploaded: {st.session_state["file_a_name"]} and {st.session_state["file_b_name"]}'
        )
    elif file_a is not None:
        st.info('Please upload File B to continue.')
    elif file_b is not None:
        st.info('Please upload File A to continue.')


def _render_step2_worksheets() -> None:
    """Step 2: Select worksheets."""
    st.subheader('Step 2: Select Worksheets')

    file_a = st.session_state.get('file_a_obj')
    file_b = st.session_state.get('file_b_obj')

    if file_a is None or file_b is None:
        st.info('Please upload both files in Step 1.')
        return

    # Load sheet names if not already cached
    if st.session_state.get('sheet_a') is None or st.session_state.get('sheet_b') is None:
        try:
            raw_a = st.session_state.get('file_a_bytes')
            raw_b = st.session_state.get('file_b_bytes')
            buffer_a = io.BytesIO(raw_a)
            buffer_b = io.BytesIO(raw_b)
            sheets_a = list_sheets(buffer_a)
            sheets_b = list_sheets(buffer_b)
            st.session_state['sheets_a'] = sheets_a
            st.session_state['sheets_b'] = sheets_b
        except Exception as exc:
            st.error(f'Failed to read workbook sheets: {exc}')
            return

    sheets_a = st.session_state.get('sheets_a', [])
    sheets_b = st.session_state.get('sheets_b', [])

    col_a, col_b = st.columns(2)

    with col_a:
        sheet_a = st.selectbox(
            'File A worksheet',
            sheets_a,
            key='selector_sheet_a',
        )
        if st.session_state.get('sheet_a') != sheet_a:
            st.session_state['sheet_a'] = sheet_a
            _reset_from_worksheet()

    with col_b:
        sheet_b = st.selectbox(
            'File B worksheet',
            sheets_b,
            key='selector_sheet_b',
        )
        if st.session_state.get('sheet_b') != sheet_b:
            st.session_state['sheet_b'] = sheet_b
            _reset_from_worksheet()

    # Load worksheets if not already loaded
    if st.session_state.get('loaded_a') is None or st.session_state.get('loaded_b') is None:
        try:
            raw_a = st.session_state.get('file_a_bytes')
            raw_b = st.session_state.get('file_b_bytes')
            buffer_a = io.BytesIO(raw_a)
            buffer_b = io.BytesIO(raw_b)
            loaded_a = load_sheet_with_metadata(buffer_a, sheet_a, track_excel_row=True)
            loaded_b = load_sheet_with_metadata(buffer_b, sheet_b, track_excel_row=True)
            st.session_state['loaded_a'] = loaded_a
            st.session_state['loaded_b'] = loaded_b
            st.session_state['columns_a'] = list(loaded_a.df.columns)
            st.session_state['columns_b'] = list(loaded_b.df.columns)
        except Exception as exc:
            st.error(f'Failed to load worksheet: {exc}')
            return

    loaded_a = st.session_state.get('loaded_a')
    loaded_b = st.session_state.get('loaded_b')
    if loaded_a is not None and loaded_b is not None:
        st.success(f'Loaded "{sheet_a}" from File A and "{sheet_b}" from File B.')


def _render_step3_keys() -> None:
    """Step 3: Configure matching keys."""
    st.subheader('Step 3: Configure Matching Keys')
    st.caption(
        'Select one or more column pairs that uniquely identify a record. '
        'Exact-name matches are suggested below; you may add or remove pairs manually.'
    )

    columns_a = st.session_state.get('columns_a', [])
    columns_b = st.session_state.get('columns_b', [])

    if not columns_a or not columns_b:
        st.info('Please select worksheets in Step 2.')
        return

    auto_pairs = discover_auto_pairs(columns_a, columns_b)

    # Current key pairs stored as stable-ID entries
    current_entries = list(st.session_state.get('key_pair_entries', []))

    # Track which entry IDs to remove (from suggestions or X buttons)
    ids_to_remove: set = set()

    # Track entry IDs added by auto-suggestion so we only remove those,
    # not manually added entries that happen to have the same columns.
    # Stored in session state so it persists across reruns.
    auto_added_ids = set(st.session_state.get('_auto_added_key_ids', set()))

    st.markdown('**Suggested key pairs (exact column names):**')
    for pair in auto_pairs:
        a = pair.file_a_column
        b = pair.file_b_column
        current_pairs = {(e['file_a'], e['file_b']) for e in current_entries}
        selected = st.checkbox(
            f'{a} ↔ {b}',
            value=(a, b) in current_pairs,
            key=f'suggest_key_{a}_{b}',
        )
        if selected and (a, b) not in current_pairs:
            # Avoid duplicate: only add if not already present
            if not any(e['file_a'] == a and e['file_b'] == b for e in current_entries):
                new_entry = {
                    'id': str(uuid.uuid4()),
                    'file_a': a,
                    'file_b': b,
                }
                current_entries.append(new_entry)
                auto_added_ids.add(new_entry['id'])
        elif not selected and (a, b) in current_pairs:
            # Only remove entries that were added by auto-suggestion,
            # not manually added entries that happen to match the same columns.
            current_entries = [
                e
                for e in current_entries
                if not (e['file_a'] == a and e['file_b'] == b and e['id'] in auto_added_ids)
            ]

    st.markdown('**Manual key mapping:**')
    updated_entries: list = []
    for entry in current_entries:
        if entry['id'] in ids_to_remove:
            continue
        cols = st.columns([2, 2, 1])
        with cols[0]:
            val_a = st.selectbox(
                'File A column',
                columns_a,
                index=None if entry['file_a'] is None else columns_a.index(entry['file_a']),
                key=f'key_pair_{entry["id"]}_a',
                placeholder='Select a column',
            )
        with cols[1]:
            val_b = st.selectbox(
                'File B column',
                columns_b,
                index=None if entry['file_b'] is None else columns_b.index(entry['file_b']),
                key=f'key_pair_{entry["id"]}_b',
                placeholder='Select a column',
            )
        with cols[2]:
            if st.button('×', key=f'remove_key_pair_{entry["id"]}'):
                ids_to_remove.add(entry['id'])
                st.session_state['key_pair_entries'] = [
                    e for e in current_entries if e['id'] not in ids_to_remove
                ]
        # Validate: no duplicate File A column across entries, no duplicate File B column.
        # Skip validation when a column is not yet selected (None).
        other_entries = [
            e for e in current_entries if e['id'] not in ids_to_remove and e['id'] != entry['id']
        ]
        other_file_a_cols = {e['file_a'] for e in other_entries if e['file_a'] is not None}
        other_file_b_cols = {e['file_b'] for e in other_entries if e['file_b'] is not None}
        duplicate_a = val_a is not None and val_a in other_file_a_cols
        duplicate_b = val_b is not None and val_b in other_file_b_cols
        if duplicate_a:
            st.error(
                f"File A column '{val_a}' is already used in another key mapping. "
                f'Choose a different column.'
            )
        if duplicate_b:
            st.error(
                f"File B column '{val_b}' is already used in another key mapping. "
                f'Choose a different column.'
            )
        if duplicate_a or duplicate_b:
            # Keep old values; do not apply the conflicting selection
            updated_entries.append(
                {
                    'id': entry['id'],
                    'file_a': entry['file_a'],
                    'file_b': entry['file_b'],
                }
            )
        else:
            # Update in place with latest widget values (None = unselected)
            updated_entries.append(
                {
                    'id': entry['id'],
                    'file_a': val_a,
                    'file_b': val_b,
                }
            )

    # Apply removals
    current_entries = [e for e in updated_entries if e['id'] not in ids_to_remove]

    if st.button('+ Add key pair', key='add_key_pair'):
        current_entries.append(
            {
                'id': str(uuid.uuid4()),
                'file_a': None,
                'file_b': None,
            }
        )
        st.session_state['key_pair_entries'] = current_entries
        st.rerun()

    # Persist current entries to session state so checkbox/selectbox
    # changes survive reruns and navigation to Step 4.
    st.session_state['key_pair_entries'] = current_entries
    # Persist auto-added IDs so unchecking suggestions only removes
    # entries that were added by auto-suggestion (not manual ones).
    st.session_state['_auto_added_key_ids'] = auto_added_ids

    if st.session_state.get('key_pair_entries'):
        current_set = {(e['file_a'], e['file_b']) for e in current_entries}
        last_key_pairs = st.session_state.get('_last_key_pairs')
        # Only reset downstream state if key mappings actually changed
        # (not on first visit when _last_key_pairs is None).
        if last_key_pairs is not None and last_key_pairs != current_set:
            _reset_from_mapping()
        st.session_state['_last_key_pairs'] = current_set

    if current_entries:
        st.success(f'{len(current_entries)} matching key pair(s) configured.')
    else:
        st.warning('Please select at least one matching key pair.')


def _render_step4_comparisons() -> None:
    """Step 4: Configure comparison columns."""
    st.subheader('Step 4: Configure Comparison Columns')
    st.caption(
        'Select which columns to compare for value changes. '
        'Matching-key columns are excluded automatically.'
    )

    columns_a = st.session_state.get('columns_a', [])
    columns_b = st.session_state.get('columns_b', [])
    key_pair_entries = st.session_state.get('key_pair_entries', [])

    if not columns_a or not columns_b:
        st.info('Please select worksheets in Step 2.')
        return

    if not key_pair_entries:
        st.info('Please configure matching keys in Step 3 first.')
        return

    key_a_cols = {e['file_a'] for e in key_pair_entries}
    key_b_cols = {e['file_b'] for e in key_pair_entries}

    available_a = [c for c in columns_a if c not in key_a_cols]
    available_b = [c for c in columns_b if c not in key_b_cols]

    # Auto-suggest exact-name compare pairs (excluding keys)
    auto_compare_pairs = discover_auto_pairs(columns_a, columns_b)
    current_entries = list(st.session_state.get('compare_pair_entries', []))
    current_set = {(e['file_a'], e['file_b']) for e in current_entries}
    current_numeric = list(st.session_state.get('numeric_pairs', []))
    current_date = list(st.session_state.get('date_pairs', []))

    # Track which entry IDs to remove
    ids_to_remove: set = set()

    # Track entry IDs added by auto-suggestion so we only remove those,
    # not manually added entries that happen to have the same columns.
    # Stored in session state so it persists across reruns.
    auto_added_ids = set(st.session_state.get('_auto_added_compare_ids', set()))

    st.markdown('**Suggested comparison pairs (exact column names):**')
    for pair in auto_compare_pairs:
        a = pair.file_a_column
        b = pair.file_b_column
        if a in key_a_cols or b in key_b_cols:
            continue
        selected = st.checkbox(
            f'{a} ↔ {b}  (Text)',
            value=(a, b) in current_set,
            key=f'suggest_compare_{a}_{b}',
        )
        if selected and (a, b) not in current_set:
            if not any(e['file_a'] == a and e['file_b'] == b for e in current_entries):
                new_entry = {
                    'id': str(uuid.uuid4()),
                    'file_a': a,
                    'file_b': b,
                    'type': 'Text',
                }
                current_entries.append(new_entry)
                auto_added_ids.add(new_entry['id'])
        elif not selected and (a, b) in current_set:
            # Only remove entries that were added by auto-suggestion,
            # not manually added entries that happen to match the same columns.
            ids_to_remove.update(
                e['id']
                for e in current_entries
                if e['file_a'] == a and e['file_b'] == b and e['id'] in auto_added_ids
            )

    st.markdown('**Manual comparison mapping:**')
    updated_entries: list = []
    for entry in current_entries:
        if entry['id'] in ids_to_remove:
            continue
        cols = st.columns([2, 2, 1, 1])
        with cols[0]:
            val_a = st.selectbox(
                'File A column',
                available_a,
                index=None if entry['file_a'] is None else available_a.index(entry['file_a']),
                key=f'compare_pair_{entry["id"]}_a',
                placeholder='Select a column',
            )
        with cols[1]:
            val_b = st.selectbox(
                'File B column',
                available_b,
                index=None if entry['file_b'] is None else available_b.index(entry['file_b']),
                key=f'compare_pair_{entry["id"]}_b',
                placeholder='Select a column',
            )
        with cols[2]:
            current_type = 'Text'
            if entry['file_a'] is not None and entry['file_b'] is not None:
                if (entry['file_a'], entry['file_b']) in current_numeric:
                    current_type = 'Numeric'
                elif (entry['file_a'], entry['file_b']) in current_date:
                    current_type = 'Date'
            comp_type = st.selectbox(
                'Type',
                ['Text', 'Numeric', 'Date'],
                index=['Text', 'Numeric', 'Date'].index(current_type),
                key=f'compare_pair_{entry["id"]}_type',
            )
        with cols[3]:
            if st.button('×', key=f'remove_compare_pair_{entry["id"]}'):
                ids_to_remove.add(entry['id'])
                st.session_state['compare_pair_entries'] = [
                    e for e in current_entries if e['id'] not in ids_to_remove
                ]
        # Validate: no duplicate File A column across comparison entries,
        # no duplicate File B column across comparison entries.
        # Skip validation when a column is not yet selected (None).
        other_entries = [
            e for e in current_entries if e['id'] not in ids_to_remove and e['id'] != entry['id']
        ]
        other_file_a_cols = {e['file_a'] for e in other_entries if e['file_a'] is not None}
        other_file_b_cols = {e['file_b'] for e in other_entries if e['file_b'] is not None}
        duplicate_a = val_a is not None and val_a in other_file_a_cols
        duplicate_b = val_b is not None and val_b in other_file_b_cols
        if duplicate_a:
            st.error(
                f"File A column '{val_a}' is already used in another comparison mapping. "
                f'Choose a different column.'
            )
        if duplicate_b:
            st.error(
                f"File B column '{val_b}' is already used in another comparison mapping. "
                f'Choose a different column.'
            )
        if duplicate_a or duplicate_b:
            # Keep old values; do not apply the conflicting selection
            updated_entries.append(
                {
                    'id': entry['id'],
                    'file_a': entry['file_a'],
                    'file_b': entry['file_b'],
                    'type': entry['type'],
                }
            )
        else:
            # Update in place with latest widget values (None = unselected)
            updated_entries.append(
                {
                    'id': entry['id'],
                    'file_a': val_a,
                    'file_b': val_b,
                    'type': comp_type,
                }
            )

    # Apply removals
    current_entries = [e for e in updated_entries if e['id'] not in ids_to_remove]

    # Rebuild type lists from updated entries (matched by entry id, not column tuples)
    new_numeric: list = []
    new_date: list = []
    for entry in current_entries:
        if entry['type'] == 'Numeric':
            new_numeric.append((entry['file_a'], entry['file_b']))
        elif entry['type'] == 'Date':
            new_date.append((entry['file_a'], entry['file_b']))
    current_numeric = new_numeric
    current_date = new_date

    if st.button('+ Add comparison pair', key='add_compare_pair'):
        if available_a and available_b:
            current_entries.append(
                {
                    'id': str(uuid.uuid4()),
                    'file_a': None,
                    'file_b': None,
                    'type': 'Text',
                }
            )
            st.session_state['compare_pair_entries'] = current_entries
            st.session_state['numeric_pairs'] = current_numeric
            st.session_state['date_pairs'] = current_date
            st.rerun()

    # Persist changes
    st.session_state['compare_pair_entries'] = current_entries
    st.session_state['numeric_pairs'] = current_numeric
    st.session_state['date_pairs'] = current_date
    st.session_state['_auto_added_compare_ids'] = auto_added_ids

    # Check if mappings changed
    current_compare_set = {(e['file_a'], e['file_b']) for e in current_entries}
    current_numeric_set = set(current_numeric)
    current_date_set = set(current_date)
    last_compare = st.session_state.get('_last_compare_pairs')
    last_numeric = st.session_state.get('_last_numeric_pairs')
    last_date = st.session_state.get('_last_date_pairs')
    if (
        last_compare != current_compare_set
        or last_numeric != current_numeric_set
        or last_date != current_date_set
    ):
        _reset_from_comparison()
        st.session_state['_last_compare_pairs'] = current_compare_set
        st.session_state['_last_numeric_pairs'] = current_numeric_set
        st.session_state['_last_date_pairs'] = current_date_set

    if current_entries:
        st.success(f'{len(current_entries)} comparison pair(s) configured.')
    else:
        st.warning(
            'Please select at least one comparison column pair. '
            'Without comparison columns, matched records will be reported as '
            '"Matched — no fields compared" rather than identical.'
        )


def _render_step5_run() -> None:
    """Step 5: Run comparison — results dashboard."""
    st.subheader('Step 5: Results')

    file_a = st.session_state.get('file_a_obj')
    file_b = st.session_state.get('file_b_obj')
    key_pair_entries = st.session_state.get('key_pair_entries', [])
    compare_pair_entries = st.session_state.get('compare_pair_entries', [])

    file_a_name = st.session_state.get('file_a_name') or 'File A'
    file_b_name = st.session_state.get('file_b_name') or 'File B'

    if file_a is None or file_b is None:
        st.info('Please upload both files.')
        return
    if not key_pair_entries:
        st.info('Please configure matching keys in Step 3.')
        return
    if not compare_pair_entries:
        st.warning(
            'No comparison columns selected. Records will be matched by key only '
            'and reported as "Matched, not compared".'
        )

    col1, col2 = st.columns([1, 4])
    with col1:
        run = st.button('Run comparison', type='primary', key='btn_run')
    with col2:
        if st.session_state.get('comparison_result') is not None:
            st.info(
                'A previous comparison result is available. Click "Run comparison" to regenerate.'
            )

    if run:
        with st.spinner('Running comparison...'):
            _run_comparison()

    result = st.session_state.get('comparison_result')
    if result is not None:
        summary = result.summary or {}
        st.subheader(
            f'Results: {file_a_name} ({summary.get("total_rows_file_a", 0)} records) '
            f'vs {file_b_name} ({summary.get("total_rows_file_b", 0)} records)'
        )

        # Dashboard metrics (never count matched-not-compared as unchanged).
        matched_records = summary.get('matched_records', 0)
        mismatched_records = summary.get('mismatched_records', 0)
        not_compared = summary.get('matched_no_fields_compared', 0)
        changed_values = sum(
            sum(1 for d in row.cell_diffs if not d.are_equal) for row in result.mismatched
        )
        duplicate_records = summary.get('duplicate_key_records', 0)
        missing_key_records = summary.get('missing_key_records', 0)
        unresolved_issues = duplicate_records + missing_key_records

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric('Records with no changes', matched_records)
            st.metric('Records with changes', mismatched_records)
        with col2:
            st.metric('Individual values changed', changed_values)
            st.metric('Records matched, not compared', not_compared)
        with col3:
            st.metric('Records missing from File B', summary.get('only_in_file_a', 0))
            st.metric('New records in File B', summary.get('only_in_file_b', 0))
        with col4:
            st.metric('Duplicate identification records', duplicate_records)
            st.metric('Records with missing identification values', missing_key_records)

        # 1. Comparison completed successfully.
        st.success('Comparison completed successfully.')

        # 2. Differences found or no differences found
        # (based only on mismatched, only_in_file_a, only_in_file_b).
        has_differences = (
            mismatched_records > 0
            or summary.get('only_in_file_a', 0) > 0
            or summary.get('only_in_file_b', 0) > 0
        )
        if has_differences:
            diff_parts = []
            if mismatched_records > 0:
                diff_parts.append(f'{mismatched_records} record(s) with changes')
            if summary.get('only_in_file_a', 0) > 0:
                diff_parts.append(f'{summary["only_in_file_a"]} record(s) missing from File B')
            if summary.get('only_in_file_b', 0) > 0:
                diff_parts.append(f'{summary["only_in_file_b"]} new record(s) in File B')
            st.info('Differences found between the files: ' + ', '.join(diff_parts) + '.')
        else:
            st.success('No differences found between the files.')

        # 3. Matching records not compared (separate informational message)
        if not_compared > 0:
            st.info(
                f'{not_compared} matched record(s) had no comparison columns '
                f'configured and were not compared.'
            )

        # 4. Data-quality problems found or none found.
        if unresolved_issues > 0:
            st.warning(
                '**What needs your attention?** '
                f'{unresolved_issues} record(s) have data-quality issues: '
                f'{duplicate_records} duplicate identification record(s) and '
                f'{missing_key_records} record(s) with missing identification '
                f'value(s). Review them on the "Records Needing Attention" tab.'
            )
        else:
            st.success('No data quality problems found.')

        _render_results_preview(result)


def _render_results_preview(result) -> None:
    """Display comparison results preview."""
    key_pair_entries = st.session_state.get('key_pair_entries', [])
    valid_key_entries = [
        e for e in key_pair_entries if e['file_a'] is not None and e['file_b'] is not None
    ]
    loaded_a = st.session_state.get('loaded_a')
    loaded_b = st.session_state.get('loaded_b')
    st.subheader('Comparison Results')

    # Tabs for each category (client-friendly sheet names).
    tab_labels = [
        'Unchanged Records',
        'Records With Changes',
        'Missing From File B',
        'New Records in File B',
        'Records Needing Attention',
    ]
    tabs = st.tabs(tab_labels)

    # Comparison columns are displayed from File B only, in configured mapping
    # order (unchanged records have identical values in both files).
    compare_pair_entries = st.session_state.get('compare_pair_entries', [])
    compare_cols_b = [e['file_b'] for e in compare_pair_entries if e['file_b'] is not None]

    # Helper to retrieve a complete original source row by DataFrame index.
    def _source_row(loaded, index_val):
        if loaded is None or index_val is None:
            return None
        try:
            return loaded.df.loc[int(index_val)]
        except (KeyError, ValueError, TypeError):
            return None

    # Unchanged Records — configured columns only:
    # interleaved keys + File B comparison columns + status + Excel row refs.
    with tabs[0]:
        rows = []
        for row_diff in result.matched:
            row_dict = {}
            # Interleaved key column values: File A key 1, File B key 1, ...
            for i, entry in enumerate(valid_key_entries):
                row_dict[file_a_column_header(entry['file_a'])] = (
                    row_diff.key_values_a[i] if i < len(row_diff.key_values_a or ()) else ''
                )
                row_dict[file_b_column_header(entry['file_b'])] = (
                    row_diff.key_values_b[i] if i < len(row_diff.key_values_b or ()) else ''
                )
            # File B comparison values (File B column names only).
            source_b = _source_row(loaded_b, row_diff.index_b)
            for col in compare_cols_b:
                row_dict[file_b_column_header(col)] = source_b[col] if source_b is not None else ''
            row_dict['Record Status'] = row_diff.status.label
            row_dict['Excel Row (File A)'] = row_diff.excel_row_a or ''
            row_dict['Excel Row (File B)'] = row_diff.excel_row_b or ''
            rows.append(row_dict)
        if rows:
            st.dataframe(rows, use_container_width=True)
        else:
            st.info('No records with no changes.')

    # Records With Changes
    with tabs[1]:
        rows = []
        for row_diff in result.mismatched:
            for cell_diff in row_diff.cell_diffs:
                if not cell_diff.are_equal:
                    row_dict = {}
                    # Interleaved key column values.
                    for i, entry in enumerate(valid_key_entries):
                        row_dict[file_a_column_header(entry['file_a'])] = (
                            row_diff.key_values_a[i] if i < len(row_diff.key_values_a or ()) else ''
                        )
                        row_dict[file_b_column_header(entry['file_b'])] = (
                            row_diff.key_values_b[i] if i < len(row_diff.key_values_b or ()) else ''
                        )
                    row_dict['Column Name (File A)'] = cell_diff.column_a
                    row_dict['Previous Value'] = cell_diff.value_a
                    row_dict['Column Name (File B)'] = cell_diff.column_b
                    row_dict['Updated Value'] = cell_diff.value_b
                    row_dict['Excel Row (File A)'] = cell_diff.excel_row_a or ''
                    row_dict['Excel Row (File B)'] = cell_diff.excel_row_b or ''
                    rows.append(row_dict)
        if rows:
            st.dataframe(rows, use_container_width=True)
        else:
            st.info('No records with changes.')

    # Missing From File B — complete original File A record (original names, no prefix).
    with tabs[2]:
        rows = []
        for row_diff in result.only_in_a:
            source_a = _source_row(loaded_a, row_diff.index_a)
            row_dict = {}
            for col in loaded_a.df.columns if loaded_a else []:
                row_dict[col] = source_a[col] if source_a is not None else ''
            row_dict['Excel Row (File A)'] = row_diff.excel_row_a or ''
            rows.append(row_dict)
        if rows:
            st.dataframe(rows, use_container_width=True)
        else:
            st.info('No records missing from File B.')

    # New Records in File B — complete original File B record (original names, no prefix).
    with tabs[3]:
        rows = []
        for row_diff in result.only_in_b:
            source_b = _source_row(loaded_b, row_diff.index_b)
            row_dict = {}
            for col in loaded_b.df.columns if loaded_b else []:
                row_dict[col] = source_b[col] if source_b is not None else ''
            row_dict['Excel Row (File B)'] = row_diff.excel_row_b or ''
            rows.append(row_dict)
        if rows:
            st.dataframe(rows, use_container_width=True)
        else:
            st.info('No new records in File B.')

    # Records Needing Attention — heading conventions applied only.
    with tabs[4]:
        rows = []
        for issue in result.data_quality_issues:
            row_dict = {}
            row_dict['Problem Found'] = issue.issue_type
            row_dict['File'] = issue.source_file
            row_dict['Excel Row'] = issue.excel_row_number or ''
            for i, entry in enumerate(valid_key_entries):
                row_dict[file_a_column_header(entry['file_a'])] = (
                    issue.key_values_a[i] if i < len(issue.key_values_a or ()) else ''
                )
                row_dict[file_b_column_header(entry['file_b'])] = (
                    issue.key_values_b[i] if i < len(issue.key_values_b or ()) else ''
                )
            row_dict['Related Column Values'] = ' | '.join(
                f'{k}={v!r}' for k, v in sorted(issue.all_column_values.items())
            )
            rows.append(row_dict)
        if rows:
            st.dataframe(rows, use_container_width=True)
        else:
            st.info('No records needing attention.')


def _render_step6_download() -> None:
    """Step 6: Download report."""
    st.subheader('Step 6: Download Report')

    report_bytes = st.session_state.get('report_bytes')
    if report_bytes is None:
        st.info('Run the comparison in Step 5 to generate a report.')
        return

    st.download_button(
        label='Download Excel Report',
        data=report_bytes,
        file_name='comparison_report.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        key='btn_download_report',
    )

    st.divider()
    st.markdown('**Finish and clear data**')
    st.caption(
        'Download the report first. Clearing removes the uploaded files, loaded '
        'worksheets, mappings, comparison results, and report data from this session.'
    )
    st.checkbox(
        'I have downloaded the report and want to clear all data.',
        key='confirm_finish_clear',
    )
    st.button(
        'Finish & Clear Data',
        key='btn_finish_clear',
        on_click=_request_clear,
        args=('confirm_finish_clear',),
    )


def main() -> None:
    st.set_page_config(
        page_title='Excel Comparison Automation',
        page_icon='\U0001f4cb',
        layout='wide',
    )

    st.title('Excel Comparison Automation')
    st.caption(
        'Compare two Excel files using configurable matching keys and '
        'generate a detailed Excel report.'
    )

    _init_session_state()

    if st.session_state.pop('_clear_blocked', False):
        st.warning('Tick the confirmation box before clearing data.')

    with st.sidebar:
        st.subheader('Session')
        st.caption('Discard the current comparison and clear all uploaded data.')
        st.checkbox('Confirm: clear all data', key='confirm_start_over')
        st.button(
            'Start Over / Clear Data',
            key='btn_start_over',
            on_click=_request_clear,
            args=('confirm_start_over',),
        )

    current_step = st.session_state.get('current_step', 1)

    # Simple step indicator
    cols = st.columns(6)
    for i, col in enumerate(cols, start=1):
        if i < current_step:
            col.markdown(f'✅ **Step {i}**')
        elif i == current_step:
            col.markdown(f'🔄 **Step {i}**')
        else:
            col.markdown(f'⬜ Step {i}')

    st.divider()

    if current_step == 1:
        _render_step1_upload()
        if (
            st.session_state.get('file_a_obj') is not None
            and st.session_state.get('file_b_obj') is not None
        ):
            if st.button('Next: Select Worksheets', key='btn_next_1'):
                st.session_state['current_step'] = 2
                st.rerun()
    elif current_step == 2:
        _render_step2_worksheets()
        col1, col2 = st.columns(2)
        with col1:
            if st.button('Back: Upload Files', key='btn_back_2'):
                st.session_state['current_step'] = 1
                st.rerun()
        with col2:
            if (
                st.session_state.get('loaded_a') is not None
                and st.session_state.get('loaded_b') is not None
            ):
                if st.button('Next: Matching Keys', key='btn_next_2'):
                    st.session_state['current_step'] = 3
                    st.rerun()
    elif current_step == 3:
        _render_step3_keys()
        col1, col2 = st.columns(2)
        with col1:
            if st.button('Back: Select Worksheets', key='btn_back_3'):
                st.session_state['current_step'] = 2
                st.rerun()
        with col2:
            if st.session_state.get('key_pair_entries'):
                if st.button('Next: Comparison Columns', key='btn_next_3'):
                    st.session_state['current_step'] = 4
                    st.rerun()
            else:
                st.button('Next: Comparison Columns', key='btn_next_3', disabled=True)
    elif current_step == 4:
        _render_step4_comparisons()
        col1, col2 = st.columns(2)
        with col1:
            if st.button('Back: Matching Keys', key='btn_back_4'):
                st.session_state['current_step'] = 3
                st.rerun()
        with col2:
            if st.button('Next: Run Comparison', key='btn_next_4'):
                st.session_state['current_step'] = 5
                st.rerun()
    elif current_step == 5:
        _render_step5_run()
        col1, col2 = st.columns(2)
        with col1:
            if st.button('Back: Comparison Columns', key='btn_back_5'):
                st.session_state['current_step'] = 4
                st.rerun()
        with col2:
            if st.session_state.get('comparison_result') is not None:
                if st.button('Next: Download Report', key='btn_next_5'):
                    st.session_state['current_step'] = 6
                    st.rerun()
    elif current_step == 6:
        _render_step6_download()
        if st.button('Back: Run Comparison', key='btn_back_6'):
            st.session_state['current_step'] = 5
            st.rerun()


if __name__ == '__main__':
    main()
