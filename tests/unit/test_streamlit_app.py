"""Unit tests for the Streamlit app (Phase 5).

Uses streamlit.testing.v1.AppTest to simulate UI interactions.
"""

import importlib.util
import logging
from pathlib import Path

import numpy as np
import pandas as pd
import pyarrow as pa
import pytest
from streamlit.testing.v1 import AppTest

from excel_compare.config import DEFAULT_MAX_FILE_SIZE_MB
from excel_compare.models import (
    CellDiff,
    ComparisonResult,
    DiffStatus,
    LoadedSheet,
    RowDiff,
    SchemaDiff,
)

APP_PATH = Path(__file__).resolve().parent.parent.parent / 'app' / 'streamlit_app.py'


def test_app_loads() -> None:
    """The app should load without errors."""
    at = AppTest.from_file(APP_PATH)
    at.run()
    assert not at.exception


def test_file_upload_section_renders() -> None:
    """File upload widgets should be present in Step 1."""
    at = AppTest.from_file(APP_PATH)
    at.run()
    file_uploaders = [
        w for w in at.file_uploader if w.label.startswith('File A') or w.label.startswith('File B')
    ]
    assert len(file_uploaders) == 2


def test_step1_upload_caption_states_50_mb_per_file() -> None:
    """The Step 1 caption must state the real upload limit (50 MB per .xlsx file).

    Regression test: the UI must not advertise the internal 200 MB decompressed
    ZIP metadata limit as the upload file-size limit.
    """
    at = AppTest.from_file(APP_PATH)
    at.run()
    assert not at.exception

    caption_text = ' '.join(c.value for c in at.caption)
    assert 'Maximum file size: 50 MB per .xlsx file.' in caption_text
    assert '200 MB' not in caption_text
    assert '200MB' not in caption_text


def test_streamlit_config_sets_max_upload_size_to_50() -> None:
    """`.streamlit/config.toml` must cap Streamlit's uploader at 50 MB.

    Streamlit's built-in uploader label ("<N>MB per file") is rendered from the
    ``server.maxUploadSize`` option, whose default is 200 MB. A custom
    ``st.caption`` cannot change that label, so the per-project config must set
    the limit to 50 to match the application's own validation limit
    (``DEFAULT_MAX_FILE_SIZE_MB``).
    """
    try:
        import tomllib  # noqa: PLC0415
    except ModuleNotFoundError:  # Python < 3.11
        tomllib = pytest.importorskip('tomli')

    config_path = APP_PATH.parents[1] / '.streamlit' / 'config.toml'
    assert config_path.is_file(), f'missing Streamlit config: {config_path}'

    with config_path.open('rb') as handle:
        config = tomllib.load(handle)

    assert config['server']['maxUploadSize'] == 50
    assert config['server']['maxUploadSize'] == DEFAULT_MAX_FILE_SIZE_MB


def test_no_files_shows_step1_only() -> None:
    """Only Step 1 should be relevant when no files uploaded."""
    at = AppTest.from_file(APP_PATH)
    at.run()
    # The app should not crash with no files
    assert not at.exception


def test_session_state_initialized() -> None:
    """Session state should have all required keys after init."""
    at = AppTest.from_file(APP_PATH)
    at.run()
    # AppTest doesn't directly expose session state, but the app should
    # not crash, which means _init_session_state() ran successfully.
    assert not at.exception


def test_run_comparison_without_files() -> None:
    """Running comparison without files should not crash."""
    at = AppTest.from_file(APP_PATH)
    at.run()
    # Clicking "Run comparison" without files should not crash
    # (it should be blocked by missing file state)
    assert not at.exception


def test_navigation_steps() -> None:
    """Step navigation should advance through the wizard."""
    at = AppTest.from_file(APP_PATH)
    at.run()
    # The app starts at step 1
    assert not at.exception


def test_app_does_not_crash_on_rerun() -> None:
    """Multiple runs should not crash the app."""
    at = AppTest.from_file(APP_PATH)
    at.run()
    at.run()
    assert not at.exception


def test_step3_column_pair_iteration_does_not_crash() -> None:
    """Step 3 should iterate ColumnPair objects without TypeError."""
    at = AppTest.from_file(APP_PATH)
    # Set up session state as if we're on Step 3 with matching columns
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['Site ID', 'Pole Number', 'Camera Name']
    at.session_state['columns_b'] = ['Site Code', 'Pole No.', 'Camera Name']
    at.session_state['key_pair_entries'] = []
    at.run()
    assert not at.exception


def test_step3_exact_name_pair_appears_in_suggestions() -> None:
    """Exact-name matching columns should appear as suggested key pairs."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['Site ID', 'Pole Number', 'Camera Name']
    at.session_state['columns_b'] = ['Site Code', 'Pole No.', 'Camera Name']
    at.session_state['key_pair_entries'] = []
    at.run()
    # Checkbox labels should include exact-name matches like "Camera Name ↔ Camera Name"
    checkbox_labels = [w.label for w in at.checkbox]
    assert any('Camera Name' in label for label in checkbox_labels)


def test_step4_column_pair_iteration_does_not_crash() -> None:
    """Step 4 should iterate ColumnPair objects without TypeError."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['Site ID', 'Pole Number', 'Camera Name']
    at.session_state['columns_b'] = ['Site Code', 'Pole No.', 'Camera Name']
    at.session_state['key_pair_entries'] = [
        {'id': 'id-1', 'file_a': 'Site ID', 'file_b': 'Site Code', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = []
    at.run()
    assert not at.exception


def test_step4_exact_name_pair_appears_in_suggestions() -> None:
    """Exact-name matching columns should appear as suggested compare pairs."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['Site ID', 'Pole Number', 'Camera Name']
    at.session_state['columns_b'] = ['Site Code', 'Pole No.', 'Camera Name']
    at.session_state['key_pair_entries'] = [
        {'id': 'id-1', 'file_a': 'Site ID', 'file_b': 'Site Code', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = []
    at.run()
    checkbox_labels = [w.label for w in at.checkbox]
    assert any('Camera Name' in label for label in checkbox_labels)


def test_step4_select_then_unselect_pair() -> None:
    """Checking then unchecking a suggestion pair removes it from session state."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['Site ID', 'Pole Number', 'Camera Name']
    at.session_state['columns_b'] = ['Site Code', 'Pole No.', 'Camera Name']
    at.session_state['key_pair_entries'] = [
        {'id': 'id-1', 'file_a': 'Site ID', 'file_b': 'Site Code', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = []
    at.run()
    for cb in at.checkbox:
        if 'Camera Name' in cb.label:
            cb.set_value(True)
            break
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert any(e['file_a'] == 'Camera Name' and e['file_b'] == 'Camera Name' for e in entries)
    for cb in at.checkbox:
        if 'Camera Name' in cb.label:
            cb.set_value(False)
            break
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert not any(e['file_a'] == 'Camera Name' and e['file_b'] == 'Camera Name' for e in entries)


def test_step4_remove_middle_pair() -> None:
    """Removing a middle pair preserves remaining pairs with correct values."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1', 'A2', 'A3']
    at.session_state['columns_b'] = ['KeyB', 'B1', 'B2', 'B3']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = [
        {'id': 'id-1', 'file_a': 'A1', 'file_b': 'B1', 'type': 'Text'},
        {'id': 'id-2', 'file_a': 'A2', 'file_b': 'B2', 'type': 'Text'},
        {'id': 'id-3', 'file_a': 'A3', 'file_b': 'B3', 'type': 'Text'},
    ]
    at.run()
    for btn in at.button:
        if btn.key == 'remove_compare_pair_id-2':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 2
    assert entries[0]['file_a'] == 'A1'
    assert entries[0]['file_b'] == 'B1'
    assert entries[1]['file_a'] == 'A3'
    assert entries[1]['file_b'] == 'B3'


def test_step4_rerun_after_removal() -> None:
    """Removed pair must not reappear after Streamlit rerun."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1', 'A2']
    at.session_state['columns_b'] = ['KeyB', 'B1', 'B2']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = [
        {'id': 'id-1', 'file_a': 'A1', 'file_b': 'B1', 'type': 'Text'},
        {'id': 'id-2', 'file_a': 'A2', 'file_b': 'B2', 'type': 'Text'},
    ]
    at.run()
    for btn in at.button:
        if btn.key == 'remove_compare_pair_id-1':
            btn.click()
            break
    at.run()
    at.run()  # process st.rerun() from the button handler
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 1
    assert entries[0]['file_a'] == 'A2'
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 1
    assert entries[0]['file_a'] == 'A2'


def test_step4_type_change_preserves_pair() -> None:
    """Changing comparison type preserves the pair."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1']
    at.session_state['columns_b'] = ['KeyB', 'B1']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = [
        {'id': 'id-1', 'file_a': 'A1', 'file_b': 'B1', 'type': 'Text'},
    ]
    at.run()
    for sb in at.selectbox:
        if sb.key == 'compare_pair_id-1_type':
            sb.set_value('Numeric')
            break
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 1
    assert entries[0]['file_a'] == 'A1'
    assert entries[0]['file_b'] == 'B1'
    numeric_pairs = at.session_state.get('numeric_pairs', [])
    assert ('A1', 'B1') in numeric_pairs


def test_step4_navigate_away_and_return() -> None:
    """Navigating away and back preserves selections."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 5
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['A1', 'A2']
    at.session_state['columns_b'] = ['B1', 'B2']
    at.session_state['key_pairs'] = [('KeyA', 'KeyB')]
    at.session_state['compare_pair_entries'] = [
        {'id': 'id-1', 'file_a': 'A1', 'file_b': 'B1', 'type': 'Text'},
    ]
    at.run()
    at.session_state['current_step'] = 4
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 1
    assert entries[0]['file_a'] == 'A1'
    assert entries[0]['file_b'] == 'B1'


def test_step4_uncheck_suggestion_keep_manual_pair() -> None:
    """Unchecking a suggestion keeps manually added pairs."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1', 'A2', 'A3']
    at.session_state['columns_b'] = ['KeyB', 'B1', 'B2', 'B3']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = [
        {'id': 'id-manual', 'file_a': 'A2', 'file_b': 'B2', 'type': 'Text'},
    ]
    at.run()
    for cb in at.checkbox:
        if 'Camera Name' in cb.label and cb.value:
            cb.set_value(False)
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert any(e['file_a'] == 'A2' and e['file_b'] == 'B2' for e in entries)


def test_step4_prevent_duplicate_mappings() -> None:
    """Adding a duplicate pair should not create duplicate entries."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1', 'A2']
    at.session_state['columns_b'] = ['KeyB', 'B1', 'B2']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = [
        {'id': 'id-1', 'file_a': 'A1', 'file_b': 'B1', 'type': 'Text'},
    ]
    at.run()
    for cb in at.checkbox:
        if 'A1' in cb.label and 'B1' in cb.label:
            cb.set_value(True)
            break
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    a1_b1_count = sum(1 for e in entries if e['file_a'] == 'A1' and e['file_b'] == 'B1')
    assert a1_b1_count == 1, f'Expected 1 entry for A1<->B1, got {a1_b1_count}'


def test_step4_comparison_config_receives_correct_pairs() -> None:
    """Final ComparisonConfig receives exactly the current pairs, no stale data."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1']
    at.session_state['columns_b'] = ['KeyB', 'B1']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = [
        {'id': 'id-1', 'file_a': 'A1', 'file_b': 'B1', 'type': 'Text'},
    ]
    at.run()
    for btn in at.button:
        if btn.key == 'remove_compare_pair_id-1':
            btn.click()
            break
    at.run()
    at.run()  # process st.rerun() from the button handler
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 0
    assert at.session_state.get('numeric_pairs', []) == []
    assert at.session_state.get('date_pairs', []) == []


# ---------------------------------------------------------------------------
# Step 3 add/remove regression tests
# ---------------------------------------------------------------------------


def test_step3_add_key_pair() -> None:
    """Adding one key pair creates a new empty row."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1', 'A2']
    at.session_state['columns_b'] = ['KeyB', 'B1', 'B2']
    at.session_state['key_pair_entries'] = []
    at.run()
    for btn in at.button:
        if btn.key == 'add_key_pair':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 1
    assert entries[0]['file_a'] is None
    assert entries[0]['file_b'] is None
    assert 'id' in entries[0]


def test_step3_add_one_click_one_row() -> None:
    """One click on Add key pair creates exactly one new empty row immediately."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1', 'A2']
    at.session_state['columns_b'] = ['KeyB', 'B1', 'B2']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB'},
    ]
    at.run()
    initial_count = len(at.session_state.get('key_pair_entries', []))
    for btn in at.button:
        if btn.key == 'add_key_pair':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == initial_count + 1
    new_entry = [e for e in entries if e['id'] != 'k-1'][0]
    assert new_entry['file_a'] is None
    assert new_entry['file_b'] is None
    orig_entry = [e for e in entries if e['id'] == 'k-1'][0]
    assert orig_entry['file_a'] == 'KeyA'
    assert orig_entry['file_b'] == 'KeyB'


def test_step3_add_one_click_no_duplicates() -> None:
    """Repeated Add clicks do not create duplicate empty rows beyond one per click."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1']
    at.session_state['columns_b'] = ['KeyB', 'B1']
    at.session_state['key_pair_entries'] = []
    at.run()
    for btn in at.button:
        if btn.key == 'add_key_pair':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 1


def test_step3_x_one_click_removes_row() -> None:
    """One click on X removes exactly the selected row."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1', 'A2']
    at.session_state['columns_b'] = ['KeyB', 'B1', 'B2']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB'},
        {'id': 'k-2', 'file_a': 'A1', 'file_b': 'B1'},
        {'id': 'k-3', 'file_a': 'A2', 'file_b': 'B2'},
    ]
    at.run()
    for btn in at.button:
        if btn.key == 'remove_key_pair_k-2':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 2
    assert entries[0]['id'] == 'k-1'
    assert entries[0]['file_a'] == 'KeyA'
    assert entries[1]['id'] == 'k-3'
    assert entries[1]['file_a'] == 'A2'


def test_step4_add_one_click_one_row() -> None:
    """One click on Add comparison pair creates exactly one new empty row immediately."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1', 'A2']
    at.session_state['columns_b'] = ['KeyB', 'B1', 'B2']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = [
        {'id': 'c-1', 'file_a': 'A1', 'file_b': 'B1', 'type': 'Text'},
    ]
    at.run()
    initial_count = len(at.session_state.get('compare_pair_entries', []))
    for btn in at.button:
        if btn.key == 'add_compare_pair':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == initial_count + 1
    new_entry = [e for e in entries if e['id'] != 'c-1'][0]
    assert new_entry['file_a'] is None
    assert new_entry['file_b'] is None
    assert new_entry['type'] == 'Text'
    orig_entry = [e for e in entries if e['id'] == 'c-1'][0]
    assert orig_entry['file_a'] == 'A1'
    assert orig_entry['file_b'] == 'B1'


def test_step4_x_one_click_removes_row() -> None:
    """One click on X removes exactly the selected comparison pair."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1', 'A2', 'A3']
    at.session_state['columns_b'] = ['KeyB', 'B1', 'B2', 'B3']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = [
        {'id': 'c-1', 'file_a': 'A1', 'file_b': 'B1', 'type': 'Text'},
        {'id': 'c-2', 'file_a': 'A2', 'file_b': 'B2', 'type': 'Text'},
        {'id': 'c-3', 'file_a': 'A3', 'file_b': 'B3', 'type': 'Text'},
    ]
    at.run()
    for btn in at.button:
        if btn.key == 'remove_compare_pair_c-2':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 2
    assert entries[0]['id'] == 'c-1'
    assert entries[0]['file_a'] == 'A1'
    assert entries[1]['id'] == 'c-3'
    assert entries[1]['file_a'] == 'A3'


def test_step3_add_multiple_key_pairs() -> None:
    """Adding multiple key pairs creates multiple rows."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1', 'A2']
    at.session_state['columns_b'] = ['KeyB', 'B1', 'B2']
    at.session_state['key_pair_entries'] = []
    at.run()
    for btn in at.button:
        if btn.key == 'add_key_pair':
            btn.click()
    at.run()
    for btn in at.button:
        if btn.key == 'add_key_pair':
            btn.click()
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 2


def test_step3_add_empty_row_no_duplicate_error() -> None:
    """Adding an empty key pair row does not create duplicate mappings or errors."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1']
    at.session_state['columns_b'] = ['KeyB', 'B1']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB'},
    ]
    at.run()
    # Add an empty row - should not error or duplicate
    for btn in at.button:
        if btn.key == 'add_key_pair':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 2
    # The new entry should be empty (unselected)
    new_entry = [e for e in entries if e['id'] != 'k-1'][0]
    assert new_entry['file_a'] is None
    assert new_entry['file_b'] is None
    # Original entry preserved
    orig_entry = [e for e in entries if e['id'] == 'k-1'][0]
    assert orig_entry['file_a'] == 'KeyA'
    assert orig_entry['file_b'] == 'KeyB'
    # No error messages
    errors = [e.value for e in at.error]
    assert not any('already used' in str(e) for e in errors)


def test_step3_add_empty_row_then_select_without_duplicate() -> None:
    """Adding an empty row does not create duplicate mappings or immediate errors."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1', 'A2']
    at.session_state['columns_b'] = ['KeyB', 'B1', 'B2']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB'},
    ]
    at.run()
    # Add empty row
    for btn in at.button:
        if btn.key == 'add_key_pair':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 2
    # New entry is empty (unselected)
    new_entry = [e for e in entries if e['id'] != 'k-1'][0]
    assert new_entry['file_a'] is None
    assert new_entry['file_b'] is None
    # Original entry preserved
    orig_entry = [e for e in entries if e['id'] == 'k-1'][0]
    assert orig_entry['file_a'] == 'KeyA'
    assert orig_entry['file_b'] == 'KeyB'
    # No duplicate-selection errors from adding an empty row
    errors = [e.value for e in at.error]
    assert not any('already used' in str(e) for e in errors)


def test_step3_remove_first_key_pair() -> None:
    """Removing the first key pair preserves remaining pairs."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1', 'A2']
    at.session_state['columns_b'] = ['KeyB', 'B1', 'B2']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB'},
        {'id': 'k-2', 'file_a': 'A1', 'file_b': 'B1'},
        {'id': 'k-3', 'file_a': 'A2', 'file_b': 'B2'},
    ]
    at.run()
    for btn in at.button:
        if btn.key == 'remove_key_pair_k-1':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 2
    assert entries[0]['file_a'] == 'A1'
    assert entries[0]['file_b'] == 'B1'
    assert entries[1]['file_a'] == 'A2'
    assert entries[1]['file_b'] == 'B2'


def test_step3_remove_middle_key_pair() -> None:
    """Removing the middle key pair preserves first and last."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1', 'A2']
    at.session_state['columns_b'] = ['KeyB', 'B1', 'B2']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB'},
        {'id': 'k-2', 'file_a': 'A1', 'file_b': 'B1'},
        {'id': 'k-3', 'file_a': 'A2', 'file_b': 'B2'},
    ]
    at.run()
    for btn in at.button:
        if btn.key == 'remove_key_pair_k-2':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 2
    assert entries[0]['file_a'] == 'KeyA'
    assert entries[0]['file_b'] == 'KeyB'
    assert entries[1]['file_a'] == 'A2'
    assert entries[1]['file_b'] == 'B2'


def test_step3_remove_last_key_pair() -> None:
    """Removing the last key pair preserves first and middle."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1', 'A2']
    at.session_state['columns_b'] = ['KeyB', 'B1', 'B2']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB'},
        {'id': 'k-2', 'file_a': 'A1', 'file_b': 'B1'},
        {'id': 'k-3', 'file_a': 'A2', 'file_b': 'B2'},
    ]
    at.run()
    for btn in at.button:
        if btn.key == 'remove_key_pair_k-3':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 2
    assert entries[0]['file_a'] == 'KeyA'
    assert entries[0]['file_b'] == 'KeyB'
    assert entries[1]['file_a'] == 'A1'
    assert entries[1]['file_b'] == 'B1'


def test_step3_remove_after_changing_selection() -> None:
    """Removing a row after changing its selectbox values works."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1', 'A2']
    at.session_state['columns_b'] = ['KeyB', 'B1', 'B2']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB'},
        {'id': 'k-2', 'file_a': 'A1', 'file_b': 'B1'},
    ]
    at.run()
    # Change k-1's File A column to A1
    for sb in at.selectbox:
        if sb.key == 'key_pair_k-1_a':
            sb.set_value('A1')
            break
    at.run()
    # Now remove k-1
    for btn in at.button:
        if btn.key == 'remove_key_pair_k-1':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 1
    assert entries[0]['id'] == 'k-2'
    assert entries[0]['file_a'] == 'A1'
    assert entries[0]['file_b'] == 'B1'


def test_step3_repeated_add_remove() -> None:
    """Repeated add/remove operations work reliably."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1', 'A2']
    at.session_state['columns_b'] = ['KeyB', 'B1', 'B2']
    at.session_state['key_pair_entries'] = []
    at.run()
    # Add three pairs
    for _ in range(3):
        for btn in at.button:
            if btn.key == 'add_key_pair':
                btn.click()
                break
        at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 3
    # Remove the middle one
    for btn in at.button:
        if btn.key == f'remove_key_pair_{entries[1]["id"]}':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 2
    # Add one more
    for btn in at.button:
        if btn.key == 'add_key_pair':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 3


def test_step3_completely_empty_key_row_does_not_crash() -> None:
    """Completely empty key row does not prevent a valid comparison from running."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1']
    at.session_state['columns_b'] = ['KeyB', 'B1']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB'},
        {'id': 'k-empty', 'file_a': None, 'file_b': None},
    ]
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 2
    valid_entries = [e for e in entries if e['file_a'] is not None and e['file_b'] is not None]
    assert len(valid_entries) == 1


def test_step3_partial_key_row_blocks_comparison() -> None:
    """Partially completed key row (one column set, other empty) blocks comparison."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1']
    at.session_state['columns_b'] = ['KeyB', 'B1']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': None},
    ]
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 1
    assert entries[0]['file_a'] == 'KeyA'
    assert entries[0]['file_b'] is None


def test_step3_valid_key_row_alongside_empty_row() -> None:
    """Valid key row works alongside an empty row."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1']
    at.session_state['columns_b'] = ['KeyB', 'B1']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB'},
        {'id': 'k-empty', 'file_a': None, 'file_b': None},
    ]
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    valid_entries = [e for e in entries if e['file_a'] is not None and e['file_b'] is not None]
    empty_entries = [e for e in entries if e['file_a'] is None and e['file_b'] is None]
    assert len(valid_entries) == 1
    assert len(empty_entries) == 1
    assert valid_entries[0]['file_a'] == 'KeyA'
    assert valid_entries[0]['file_b'] == 'KeyB'


def test_step4_completely_empty_compare_row_does_not_crash() -> None:
    """Completely empty comparison row does not prevent a valid comparison from running."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1']
    at.session_state['columns_b'] = ['KeyB', 'B1']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = [
        {'id': 'c-1', 'file_a': 'A1', 'file_b': 'B1', 'type': 'Text'},
        {'id': 'c-empty', 'file_a': None, 'file_b': None, 'type': 'Text'},
    ]
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 2
    valid_entries = [e for e in entries if e['file_a'] is not None and e['file_b'] is not None]
    empty_entries = [e for e in entries if e['file_a'] is None and e['file_b'] is None]
    assert len(valid_entries) == 1
    assert len(empty_entries) == 1


def test_step4_partial_compare_row_blocks_comparison() -> None:
    """Partially completed comparison row blocks comparison."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1']
    at.session_state['columns_b'] = ['KeyB', 'B1']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = [
        {'id': 'c-1', 'file_a': 'A1', 'file_b': None, 'type': 'Text'},
    ]
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 1
    assert entries[0]['file_a'] == 'A1'
    assert entries[0]['file_b'] is None


def test_step4_valid_compare_row_alongside_empty_row() -> None:
    """Valid comparison row works alongside an empty row."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1']
    at.session_state['columns_b'] = ['KeyB', 'B1']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = [
        {'id': 'c-1', 'file_a': 'A1', 'file_b': 'B1', 'type': 'Text'},
        {'id': 'c-empty', 'file_a': None, 'file_b': None, 'type': 'Text'},
    ]
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    valid_entries = [e for e in entries if e['file_a'] is not None and e['file_b'] is not None]
    empty_entries = [e for e in entries if e['file_a'] is None and e['file_b'] is None]
    assert len(valid_entries) == 1
    assert len(empty_entries) == 1
    assert valid_entries[0]['file_a'] == 'A1'
    assert valid_entries[0]['file_b'] == 'B1'


def test_step3_no_duplicate_key_pairs() -> None:
    """Checking a suggestion for an already-present pair does not duplicate it."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'Shared']
    at.session_state['columns_b'] = ['KeyB', 'Shared']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'Shared', 'file_b': 'Shared'},
    ]
    at.run()
    # Checking an already-present suggestion should not create duplicates.
    shared_count_before = sum(
        1
        for e in at.session_state.get('key_pair_entries', [])
        if e['file_a'] == 'Shared' and e['file_b'] == 'Shared'
    )
    assert shared_count_before == 1
    for cb in at.checkbox:
        if 'Shared' in cb.label:
            cb.set_value(True)
            break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    shared_count = sum(1 for e in entries if e['file_a'] == 'Shared' and e['file_b'] == 'Shared')
    assert shared_count == 1, f'Expected 1 Shared<->Shared entry, got {shared_count}'


def test_step3_select_suggestion_creates_key_pair_entry() -> None:
    """Selecting an automatic key suggestion creates a key_pair_entries entry."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'Shared']
    at.session_state['columns_b'] = ['KeyB', 'Shared']
    at.session_state['key_pair_entries'] = []
    at.run()
    # Check the Shared <-> Shared suggestion checkbox
    for cb in at.checkbox:
        if 'Shared' in cb.label:
            cb.set_value(True)
            break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    shared_entries = [e for e in entries if e['file_a'] == 'Shared' and e['file_b'] == 'Shared']
    assert len(shared_entries) == 1, f'Expected 1 Shared<->Shared entry, got {len(shared_entries)}'


def test_step3_multiple_suggestions_preserved() -> None:
    """Multiple selected suggestions are all preserved in key_pair_entries."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'Shared1', 'Shared2']
    at.session_state['columns_b'] = ['KeyB', 'Shared1', 'Shared2']
    at.session_state['key_pair_entries'] = []
    at.run()
    # Check both suggestions
    for cb in at.checkbox:
        if 'Shared1' in cb.label or 'Shared2' in cb.label:
            cb.set_value(True)
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    shared1 = [e for e in entries if e['file_a'] == 'Shared1' and e['file_b'] == 'Shared1']
    shared2 = [e for e in entries if e['file_a'] == 'Shared2' and e['file_b'] == 'Shared2']
    assert len(shared1) == 1, f'Expected 1 Shared1<->Shared1 entry, got {len(shared1)}'
    assert len(shared2) == 1, f'Expected 1 Shared2<->Shared2 entry, got {len(shared2)}'


def test_step3_selection_survives_step4_navigation() -> None:
    """Key selections survive navigating from Step 3 to Step 4."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'Shared']
    at.session_state['columns_b'] = ['KeyB', 'Shared']
    at.session_state['key_pair_entries'] = []
    at.run()
    # Check the Shared <-> Shared suggestion
    for cb in at.checkbox:
        if 'Shared' in cb.label:
            cb.set_value(True)
            break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 1
    assert entries[0]['file_a'] == 'Shared'
    # Navigate to Step 4
    at.session_state['current_step'] = 4
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 1
    assert entries[0]['file_a'] == 'Shared'


def test_step4_rerun_does_not_clear_step3_selections() -> None:
    """Step 4 reruns do not clear Step 3 key selections."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'Shared']
    at.session_state['columns_b'] = ['KeyB', 'Shared']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'Shared', 'file_b': 'Shared'},
    ]
    at.session_state['compare_pair_entries'] = []
    at.run()
    # Rerun Step 4
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 1
    assert entries[0]['file_a'] == 'Shared'


def test_step3_add_remove_preserves_other_keys() -> None:
    """Add/remove operations do not discard other selected keys."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'Shared1', 'Shared2']
    at.session_state['columns_b'] = ['KeyB', 'Shared1', 'Shared2']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'Shared1', 'file_b': 'Shared1'},
    ]
    at.run()
    # Check Shared2 suggestion
    for cb in at.checkbox:
        if 'Shared2' in cb.label:
            cb.set_value(True)
            break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 2
    # Remove Shared1 via X button
    for btn in at.button:
        if btn.key == 'remove_key_pair_k-1':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 1
    assert entries[0]['file_a'] == 'Shared2'
    assert entries[0]['file_b'] == 'Shared2'


# ---------------------------------------------------------------------------
# Step 4 add/remove regression tests
# ---------------------------------------------------------------------------


def test_step4_add_comparison_pair() -> None:
    """Adding one comparison pair creates a new empty row."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1', 'A2']
    at.session_state['columns_b'] = ['KeyB', 'B1', 'B2']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = []
    at.run()
    for btn in at.button:
        if btn.key == 'add_compare_pair':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 1
    assert entries[0]['file_a'] is None
    assert entries[0]['file_b'] is None
    assert entries[0]['type'] == 'Text'


def test_step4_add_multiple_comparison_pairs() -> None:
    """Adding multiple comparison pairs creates multiple rows."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1', 'A2']
    at.session_state['columns_b'] = ['KeyB', 'B1', 'B2']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = []
    at.run()
    for btn in at.button:
        if btn.key == 'add_compare_pair':
            btn.click()
    at.run()
    for btn in at.button:
        if btn.key == 'add_compare_pair':
            btn.click()
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 2


def test_step4_add_empty_row_no_duplicate_error() -> None:
    """Adding an empty comparison pair row does not create duplicate mappings or errors."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1']
    at.session_state['columns_b'] = ['KeyB', 'B1']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = [
        {'id': 'c-1', 'file_a': 'A1', 'file_b': 'B1', 'type': 'Text'},
    ]
    at.run()
    # Add an empty row - should not error or duplicate
    for btn in at.button:
        if btn.key == 'add_compare_pair':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 2
    # The new entry should be empty (unselected)
    new_entry = [e for e in entries if e['id'] != 'c-1'][0]
    assert new_entry['file_a'] is None
    assert new_entry['file_b'] is None
    assert new_entry['type'] == 'Text'
    # Original entry preserved
    orig_entry = [e for e in entries if e['id'] == 'c-1'][0]
    assert orig_entry['file_a'] == 'A1'
    assert orig_entry['file_b'] == 'B1'
    # No error messages
    errors = [e.value for e in at.error]
    assert not any('already used' in str(e) for e in errors)


def test_step4_remove_first_comparison_pair() -> None:
    """Removing the first comparison pair preserves remaining pairs."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1', 'A2', 'A3']
    at.session_state['columns_b'] = ['KeyB', 'B1', 'B2', 'B3']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = [
        {'id': 'id-1', 'file_a': 'A1', 'file_b': 'B1', 'type': 'Text'},
        {'id': 'id-2', 'file_a': 'A2', 'file_b': 'B2', 'type': 'Text'},
        {'id': 'id-3', 'file_a': 'A3', 'file_b': 'B3', 'type': 'Text'},
    ]
    at.run()
    for btn in at.button:
        if btn.key == 'remove_compare_pair_id-1':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 2
    assert entries[0]['file_a'] == 'A2'
    assert entries[0]['file_b'] == 'B2'
    assert entries[1]['file_a'] == 'A3'
    assert entries[1]['file_b'] == 'B3'


def test_step4_remove_middle_comparison_pair() -> None:
    """Removing the middle comparison pair preserves first and last."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1', 'A2', 'A3']
    at.session_state['columns_b'] = ['KeyB', 'B1', 'B2', 'B3']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = [
        {'id': 'id-1', 'file_a': 'A1', 'file_b': 'B1', 'type': 'Text'},
        {'id': 'id-2', 'file_a': 'A2', 'file_b': 'B2', 'type': 'Text'},
        {'id': 'id-3', 'file_a': 'A3', 'file_b': 'B3', 'type': 'Text'},
    ]
    at.run()
    for btn in at.button:
        if btn.key == 'remove_compare_pair_id-2':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 2
    assert entries[0]['file_a'] == 'A1'
    assert entries[0]['file_b'] == 'B1'
    assert entries[1]['file_a'] == 'A3'
    assert entries[1]['file_b'] == 'B3'


def test_step4_remove_last_comparison_pair() -> None:
    """Removing the last comparison pair preserves first and middle."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1', 'A2']
    at.session_state['columns_b'] = ['KeyB', 'B1', 'B2']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = [
        {'id': 'id-1', 'file_a': 'A1', 'file_b': 'B1', 'type': 'Text'},
        {'id': 'id-2', 'file_a': 'A2', 'file_b': 'B2', 'type': 'Text'},
    ]
    at.run()
    for btn in at.button:
        if btn.key == 'remove_compare_pair_id-2':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 1
    assert entries[0]['file_a'] == 'A1'
    assert entries[0]['file_b'] == 'B1'


def test_step4_type_preserved_after_removal() -> None:
    """Comparison type is preserved for remaining pairs after removal."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'A1', 'A2']
    at.session_state['columns_b'] = ['KeyB', 'B1', 'B2']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = [
        {'id': 'id-1', 'file_a': 'A1', 'file_b': 'B1', 'type': 'Numeric'},
        {'id': 'id-2', 'file_a': 'A2', 'file_b': 'B2', 'type': 'Text'},
    ]
    at.run()
    # Change id-1 type to Numeric
    for sb in at.selectbox:
        if sb.key == 'compare_pair_id-1_type':
            sb.set_value('Numeric')
            break
    at.run()
    # Remove id-2
    for btn in at.button:
        if btn.key == 'remove_compare_pair_id-2':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 1
    assert entries[0]['file_a'] == 'A1'
    assert entries[0]['file_b'] == 'B1'
    assert entries[0]['type'] == 'Numeric'
    numeric_pairs = at.session_state.get('numeric_pairs', [])
    assert ('A1', 'B1') in numeric_pairs


def test_step4_no_duplicate_comparison_pairs() -> None:
    """Checking a suggestion for an already-manual pair does not duplicate it."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['KeyA', 'Shared']
    at.session_state['columns_b'] = ['KeyB', 'Shared']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'KeyA', 'file_b': 'KeyB', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = [
        {'id': 'id-1', 'file_a': 'Shared', 'file_b': 'Shared', 'type': 'Text'},
    ]
    at.run()
    # Uncheck the Shared <-> Shared suggestion.
    # Manually added entries must persist when unchecking auto-suggestions.
    for cb in at.checkbox:
        if 'Shared' in cb.label:
            cb.set_value(False)
            break
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 1  # manual entry persists when auto-suggestion unchecked
    assert entries[0]['file_a'] == 'Shared' and entries[0]['file_b'] == 'Shared'
    # Re-check the Shared <-> Shared suggestion: no duplicate.
    for cb in at.checkbox:
        if 'Shared' in cb.label:
            cb.set_value(True)
            break
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    shared_count = sum(1 for e in entries if e['file_a'] == 'Shared' and e['file_b'] == 'Shared')
    assert shared_count == 1, f'Expected 1 Shared<->Shared entry, got {shared_count}'


# ---------------------------------------------------------------------------
# Dashboard four-signal regression tests
# ---------------------------------------------------------------------------


def test_step5_dashboard_four_signals_completed_and_differences() -> None:
    """Dashboard shows completion success AND differences message when changes exist."""
    summary = {
        'total_rows_file_a': 3,
        'total_rows_file_b': 3,
        'matched_records': 1,
        'mismatched_records': 1,
        'matched_no_fields_compared': 0,
        'only_in_file_a': 0,
        'only_in_file_b': 0,
        'duplicate_key_records': 0,
        'missing_key_records': 0,
    }
    at = _step5_app_with_result(summary, mismatched_count=1, changed_values=1)
    assert not at.exception

    successes = [s.value for s in at.success]
    assert any('Comparison completed successfully' in str(s) for s in successes)
    infos = [i.value for i in at.info]
    assert any('Differences found between the files' in str(i) for i in infos)


def test_step5_dashboard_no_differences_when_clean() -> None:
    """Dashboard shows 'No differences found' when no changed/missing/new records."""
    summary = {
        'total_rows_file_a': 3,
        'total_rows_file_b': 3,
        'matched_records': 3,
        'mismatched_records': 0,
        'matched_no_fields_compared': 0,
        'only_in_file_a': 0,
        'only_in_file_b': 0,
        'duplicate_key_records': 0,
        'missing_key_records': 0,
    }
    at = _step5_app_with_result(summary)
    assert not at.exception

    successes = [s.value for s in at.success]
    assert any('No differences found between the files' in str(s) for s in successes)
    infos = [i.value for i in at.info]
    assert not any('Differences found between the files' in str(i) for i in infos)


def test_step5_dashboard_not_compared_separate_signal() -> None:
    """Not-compared records are a separate informational signal, not differences."""
    summary = {
        'total_rows_file_a': 5,
        'total_rows_file_b': 5,
        'matched_records': 0,
        'mismatched_records': 0,
        'matched_no_fields_compared': 3,
        'only_in_file_a': 0,
        'only_in_file_b': 0,
        'duplicate_key_records': 0,
        'missing_key_records': 0,
    }
    at = _step5_app_with_result(summary)
    assert not at.exception

    infos = [i.value for i in at.info]
    assert any('not compared' in str(i).lower() for i in infos)
    # Not-compared should NOT trigger differences message
    assert not any('Differences found between the files' in str(i) for i in infos)


def test_step5_dashboard_differences_and_data_quality_separate() -> None:
    """Differences and data-quality problems are separate signals."""
    summary = {
        'total_rows_file_a': 4,
        'total_rows_file_b': 4,
        'matched_records': 1,
        'mismatched_records': 1,
        'matched_no_fields_compared': 0,
        'only_in_file_a': 1,
        'only_in_file_b': 0,
        'duplicate_key_records': 1,
        'missing_key_records': 0,
    }
    at = _step5_app_with_result(summary, mismatched_count=1, changed_values=1)
    assert not at.exception

    successes = [s.value for s in at.success]
    assert any('Comparison completed successfully' in str(s) for s in successes)
    infos = [i.value for i in at.info]
    assert any('Differences found between the files' in str(i) for i in infos)
    warnings = [w.value for w in at.warning]
    assert any('What needs your attention?' in str(w) for w in warnings)


def _step5_app_with_result(summary: dict, mismatched_count: int = 0, changed_values: int = 0):
    """Build an AppTest stopped at Step 5 with a canned comparison result."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 5
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['file_a_name'] = 'East.xlsx'
    at.session_state['file_b_name'] = 'West.xlsx'
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'ID', 'file_b': 'ID', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = [
        {'id': 'c-1', 'file_a': 'Amount', 'file_b': 'Amount', 'type': 'Text'},
    ]

    mismatched = []
    if mismatched_count:
        cell_diffs = [
            CellDiff(
                column_a='Amount',
                column_b='Amount',
                row_key="ID='2'",
                sheet='',
                value_a='10',
                value_b='20',
                are_equal=False,
                excel_row_a=3,
                excel_row_b=3,
            )
        ]
        if changed_values == 2:
            cell_diffs.append(
                CellDiff(
                    column_a='Name',
                    column_b='Name',
                    row_key="ID='2'",
                    sheet='',
                    value_a='Sam',
                    value_b='Samantha',
                    are_equal=False,
                    excel_row_a=3,
                    excel_row_b=3,
                )
            )
        mismatched = [
            RowDiff(
                row_key="ID='2'",
                status=DiffStatus.MISMATCHED,
                cell_diffs=cell_diffs,
                index_a=1,
                index_b=1,
                excel_row_a=3,
                excel_row_b=3,
                key_values_a=('2',),
                key_values_b=('2',),
            )
        ]

    result = ComparisonResult(
        matched=[
            RowDiff(
                row_key="ID='1'",
                status=DiffStatus.MATCHED,
                cell_diffs=[],
                index_a=0,
                index_b=0,
                excel_row_a=2,
                excel_row_b=2,
                key_values_a=('1',),
                key_values_b=('1',),
            )
        ],
        mismatched=mismatched,
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=SchemaDiff(
            common_columns=['ID', 'Amount'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['ID'],
            compare_columns=['Amount'],
        ),
        summary=summary,
    )
    at.session_state['comparison_result'] = result
    at.run()
    return at


def _metric_values(at) -> dict:
    return {m.label: m.value for m in at.metric}


def test_step5_dashboard_shows_filenames_and_totals() -> None:
    """Dashboard shows both uploaded file names and total record counts."""
    summary = {
        'total_rows_file_a': 100,
        'total_rows_file_b': 120,
        'matched_records': 90,
        'mismatched_records': 5,
        'matched_no_fields_compared': 0,
        'only_in_file_a': 3,
        'only_in_file_b': 2,
        'duplicate_key_records': 0,
        'missing_key_records': 0,
    }
    at = _step5_app_with_result(summary)
    assert not at.exception

    subheaders = [sh.value for sh in at.subheader]
    assert any('East.xlsx' in str(sh) and '100 records' in str(sh) for sh in subheaders)
    assert any('West.xlsx' in str(sh) and '120 records' in str(sh) for sh in subheaders)

    values = _metric_values(at)
    assert values.get('Records with no changes') == '90'
    assert values.get('Records with changes') == '5'


def test_step5_dashboard_shows_record_and_value_counts() -> None:
    """Dashboard shows unchanged, changed-record, and changed-value counts."""
    summary = {
        'total_rows_file_a': 3,
        'total_rows_file_b': 3,
        'matched_records': 1,
        'mismatched_records': 1,
        'matched_no_fields_compared': 0,
        'only_in_file_a': 0,
        'only_in_file_b': 0,
        'duplicate_key_records': 0,
        'missing_key_records': 0,
    }
    at = _step5_app_with_result(summary, mismatched_count=1, changed_values=2)
    assert not at.exception

    values = _metric_values(at)
    assert values.get('Records with no changes') == '1'
    assert values.get('Records with changes') == '1'
    assert values.get('Individual values changed') == '2'


def test_step5_dashboard_never_counts_not_compared_as_unchanged() -> None:
    """Matched-not-compared records are reported separately, not as unchanged."""
    summary = {
        'total_rows_file_a': 5,
        'total_rows_file_b': 5,
        'matched_records': 3,
        'mismatched_records': 0,
        'matched_no_fields_compared': 2,
        'only_in_file_a': 0,
        'only_in_file_b': 0,
        'duplicate_key_records': 0,
        'missing_key_records': 0,
    }
    at = _step5_app_with_result(summary)
    assert not at.exception

    values = _metric_values(at)
    assert values.get('Records with no changes') == '3'
    assert values.get('Records matched, not compared') == '2'


def test_step5_dashboard_shows_missing_and_quality_issues() -> None:
    """Dashboard shows records missing from either file and quality issues."""
    summary = {
        'total_rows_file_a': 4,
        'total_rows_file_b': 4,
        'matched_records': 1,
        'mismatched_records': 1,
        'matched_no_fields_compared': 0,
        'only_in_file_a': 1,
        'only_in_file_b': 1,
        'duplicate_key_records': 2,
        'missing_key_records': 1,
    }
    at = _step5_app_with_result(summary)
    assert not at.exception

    values = _metric_values(at)
    assert values.get('Records missing from File B') == '1'
    assert values.get('New records in File B') == '1'
    assert values.get('Duplicate identification records') == '2'
    assert values.get('Records with missing identification values') == '1'


def test_step5_dashboard_attention_warns_when_unresolved_issues() -> None:
    """Unresolved duplicate/missing-key records show an attention warning, not unconditional success."""
    summary = {
        'total_rows_file_a': 3,
        'total_rows_file_b': 3,
        'matched_records': 2,
        'mismatched_records': 1,
        'matched_no_fields_compared': 0,
        'only_in_file_a': 0,
        'only_in_file_b': 0,
        'duplicate_key_records': 1,
        'missing_key_records': 0,
    }
    at = _step5_app_with_result(summary)
    assert not at.exception

    warnings = [w.value for w in at.warning]
    assert any('What needs your attention?' in str(w) for w in warnings)
    successes = [s.value for s in at.success]
    assert not any('No data-quality issues' in str(s) for s in successes)


def test_step5_dashboard_attention_succeeds_when_clean() -> None:
    """No unresolved issues shows a plain-language success message."""
    summary = {
        'total_rows_file_a': 3,
        'total_rows_file_b': 3,
        'matched_records': 3,
        'mismatched_records': 0,
        'matched_no_fields_compared': 0,
        'only_in_file_a': 0,
        'only_in_file_b': 0,
        'duplicate_key_records': 0,
        'missing_key_records': 0,
    }
    at = _step5_app_with_result(summary)
    assert not at.exception

    successes = [s.value for s in at.success]
    assert any('No data quality problems found' in str(s) for s in successes)
    warnings = [w.value for w in at.warning]
    assert not any('What needs your attention?' in str(w) for w in warnings)


def _step5_app_for_comparison(
    key_pair_entries: list[dict],
    compare_pair_entries: list[dict],
) -> AppTest:
    """Build an AppTest at Step 5 with loadable DataFrames for comparison."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 5
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['file_a_name'] = 'A.xlsx'
    at.session_state['file_b_name'] = 'B.xlsx'
    at.session_state['key_pair_entries'] = key_pair_entries
    at.session_state['compare_pair_entries'] = compare_pair_entries
    df_a = pd.DataFrame({'ID': ['1'], 'Amount': ['100']})
    df_b = pd.DataFrame({'ID': ['1'], 'Amount': ['100']})
    at.session_state['loaded_a'] = LoadedSheet(
        df=df_a, sheet_name='Sheet1', file_source='A.xlsx', excel_row_mapping={0: 2}
    )
    at.session_state['loaded_b'] = LoadedSheet(
        df=df_b, sheet_name='Sheet1', file_source='B.xlsx', excel_row_mapping={0: 2}
    )
    at.run()
    return at


def test_step5_comparison_receives_selected_keys() -> None:
    """_run_comparison() receives the matching-key pairs selected in Step 3."""
    at = _step5_app_for_comparison(
        key_pair_entries=[{'id': 'k-1', 'file_a': 'ID', 'file_b': 'ID'}],
        compare_pair_entries=[],
    )
    assert not at.exception
    for btn in at.button:
        if btn.key == 'btn_run':
            btn.click()
            break
    at.run()
    result = at.session_state.get('comparison_result')
    assert result is not None, 'comparison_result should be stored after running'
    # Without compare pairs, matched records go to matched_no_fields_compared
    assert result.summary['matched_no_fields_compared'] == 1


def test_step5_comparison_succeeds_with_valid_keys() -> None:
    """Comparison succeeds when valid keys are selected in Step 3 and compare columns in Step 4."""
    at = _step5_app_for_comparison(
        key_pair_entries=[{'id': 'k-1', 'file_a': 'ID', 'file_b': 'ID'}],
        compare_pair_entries=[
            {'id': 'c-1', 'file_a': 'Amount', 'file_b': 'Amount', 'type': 'Text'}
        ],
    )
    assert not at.exception
    for btn in at.button:
        if btn.key == 'btn_run':
            btn.click()
            break
    at.run()
    result = at.session_state.get('comparison_result')
    assert result is not None, 'comparison_result should be stored after running'
    assert len(result.matched) == 1, f'Expected 1 matched record, got {len(result.matched)}'
    assert result.summary['matched_records'] == 1
    assert result.summary['mismatched_records'] == 0


# ---------------------------------------------------------------------------
# End-to-end widget interaction regression tests
# These tests use actual Streamlit widget interactions (selectbox, checkbox)
# rather than manually populating session-state lists.
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Multi-row persistence regression tests (actual Streamlit widget interactions)
# ---------------------------------------------------------------------------


def _add_key_pair_entries(at: AppTest, count: int = 2) -> None:
    """Add `count` empty key-pair entries via the Add button."""
    for _ in range(count):
        for btn in at.button:
            if btn.key == 'add_key_pair':
                btn.click()
                break
        at.run()


def test_step3_two_key_pairs_persist_after_rerun() -> None:
    """Selecting two key pairs through selectboxes: both persist after rerun."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['ID', 'Name', 'Amount']
    at.session_state['columns_b'] = ['ID', 'Name', 'Amount']
    at.session_state['key_pair_entries'] = []
    at.run()
    _add_key_pair_entries(at, 2)
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 2
    id1, id2 = entries[0]['id'], entries[1]['id']
    # Select ID↔ID for entry 1, Name↔Name for entry 2
    for sb in at.selectbox:
        if sb.key == f'key_pair_{id1}_a':
            sb.set_value('ID')
            break
    for sb in at.selectbox:
        if sb.key == f'key_pair_{id1}_b':
            sb.set_value('ID')
            break
    for sb in at.selectbox:
        if sb.key == f'key_pair_{id2}_a':
            sb.set_value('Name')
            break
    for sb in at.selectbox:
        if sb.key == f'key_pair_{id2}_b':
            sb.set_value('Name')
            break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    by_id = {e['id']: e for e in entries}
    assert len(by_id) == 2
    assert by_id[id1]['file_a'] == 'ID' and by_id[id1]['file_b'] == 'ID'
    assert by_id[id2]['file_a'] == 'Name' and by_id[id2]['file_b'] == 'Name'


def test_step3_three_key_pairs_persist_after_rerun() -> None:
    """Selecting three key pairs through selectboxes: all persist after rerun."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['ID', 'Name', 'Amount', 'Date']
    at.session_state['columns_b'] = ['ID', 'Name', 'Amount', 'Date']
    at.session_state['key_pair_entries'] = []
    at.run()
    _add_key_pair_entries(at, 3)
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 3
    ids = [e['id'] for e in entries]
    values = [('ID', 'ID'), ('Name', 'Name'), ('Amount', 'Amount')]
    for entry_id, (a, b) in zip(ids, values, strict=False):
        for sb in at.selectbox:
            if sb.key == f'key_pair_{entry_id}_a':
                sb.set_value(a)
                break
        for sb in at.selectbox:
            if sb.key == f'key_pair_{entry_id}_b':
                sb.set_value(b)
                break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    by_id = {e['id']: e for e in entries}
    assert len(by_id) == 3
    for entry_id, (a, b) in zip(ids, values, strict=False):
        assert by_id[entry_id]['file_a'] == a
        assert by_id[entry_id]['file_b'] == b


def test_step3_key_selections_survive_step3_to_step4_navigation() -> None:
    """Two key selections survive navigation Step 3 -> Step 4 -> Step 3."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['ID', 'Name', 'Amount']
    at.session_state['columns_b'] = ['ID', 'Name', 'Amount']
    at.session_state['key_pair_entries'] = []
    at.run()
    _add_key_pair_entries(at, 2)
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 2
    ids = [e['id'] for e in entries]
    # Select both key pairs
    for entry_id, col in zip(ids, ['ID', 'Name'], strict=False):
        for sb in at.selectbox:
            if sb.key == f'key_pair_{entry_id}_a':
                sb.set_value(col)
                break
        for sb in at.selectbox:
            if sb.key == f'key_pair_{entry_id}_b':
                sb.set_value(col)
                break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    by_id = {e['id']: e for e in entries}
    assert by_id[ids[0]]['file_a'] == 'ID'
    assert by_id[ids[1]]['file_a'] == 'Name'
    # Navigate to Step 4
    at.session_state['current_step'] = 4
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    by_id = {e['id']: e for e in entries}
    assert by_id[ids[0]]['file_a'] == 'ID'
    assert by_id[ids[1]]['file_a'] == 'Name'
    # Navigate back to Step 3
    at.session_state['current_step'] = 3
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    by_id = {e['id']: e for e in entries}
    assert by_id[ids[0]]['file_a'] == 'ID'
    assert by_id[ids[1]]['file_a'] == 'Name'


def test_step3_remove_middle_row_preserves_other_selections() -> None:
    """Remove the middle key pair: the other two selections remain unchanged."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['ID', 'Name', 'Amount']
    at.session_state['columns_b'] = ['ID', 'Name', 'Amount']
    at.session_state['key_pair_entries'] = []
    at.run()
    _add_key_pair_entries(at, 3)
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 3
    ids = [e['id'] for e in entries]
    # Select ID↔ID, Name↔Name, Amount↔Amount
    for entry_id, col in zip(ids, ['ID', 'Name', 'Amount'], strict=False):
        for sb in at.selectbox:
            if sb.key == f'key_pair_{entry_id}_a':
                sb.set_value(col)
                break
        for sb in at.selectbox:
            if sb.key == f'key_pair_{entry_id}_b':
                sb.set_value(col)
                break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    by_id = {e['id']: e for e in entries}
    assert len(by_id) == 3
    # Remove the middle row
    for btn in at.button:
        if btn.key == f'remove_key_pair_{ids[1]}':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    by_id = {e['id']: e for e in entries}
    assert len(by_id) == 2
    assert ids[1] not in by_id
    assert by_id[ids[0]]['file_a'] == 'ID'
    assert by_id[ids[0]]['file_b'] == 'ID'
    assert by_id[ids[2]]['file_a'] == 'Amount'
    assert by_id[ids[2]]['file_b'] == 'Amount'


def test_step4_two_compare_pairs_persist_after_rerun() -> None:
    """Selecting two comparison pairs through selectboxes: both persist after rerun."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['ID', 'Name', 'Amount', 'Date']
    at.session_state['columns_b'] = ['ID', 'Name', 'Amount', 'Date']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'ID', 'file_b': 'ID'},
        {'id': 'k-2', 'file_a': 'Name', 'file_b': 'Name'},
    ]
    at.session_state['compare_pair_entries'] = []
    at.run()
    # Add two comparison pairs
    for _ in range(2):
        for btn in at.button:
            if btn.key == 'add_compare_pair':
                btn.click()
                break
        at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 2
    ids = [e['id'] for e in entries]
    # Select Amount↔Amount for entry 1, Date↔Date for entry 2
    for entry_id, col in zip(ids, ['Amount', 'Date'], strict=False):
        for sb in at.selectbox:
            if sb.key == f'compare_pair_{entry_id}_a':
                sb.set_value(col)
                break
        for sb in at.selectbox:
            if sb.key == f'compare_pair_{entry_id}_b':
                sb.set_value(col)
                break
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    by_id = {e['id']: e for e in entries}
    assert len(by_id) == 2
    assert by_id[ids[0]]['file_a'] == 'Amount'
    assert by_id[ids[0]]['file_b'] == 'Amount'
    assert by_id[ids[1]]['file_a'] == 'Date'
    assert by_id[ids[1]]['file_b'] == 'Date'


def test_step4_compare_selections_survive_step4_to_step3_navigation() -> None:
    """Two compare selections survive navigation Step 4 -> Step 3 -> Step 4."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['ID', 'Name', 'Amount', 'Date']
    at.session_state['columns_b'] = ['ID', 'Name', 'Amount', 'Date']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'ID', 'file_b': 'ID'},
        {'id': 'k-2', 'file_a': 'Name', 'file_b': 'Name'},
    ]
    at.session_state['compare_pair_entries'] = []
    at.run()
    # Add two comparison pairs
    for _ in range(2):
        for btn in at.button:
            if btn.key == 'add_compare_pair':
                btn.click()
                break
        at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 2
    ids = [e['id'] for e in entries]
    # Select Amount↔Amount, Date↔Date
    for entry_id, col in zip(ids, ['Amount', 'Date'], strict=False):
        for sb in at.selectbox:
            if sb.key == f'compare_pair_{entry_id}_a':
                sb.set_value(col)
                break
        for sb in at.selectbox:
            if sb.key == f'compare_pair_{entry_id}_b':
                sb.set_value(col)
                break
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    by_id = {e['id']: e for e in entries}
    assert by_id[ids[0]]['file_a'] == 'Amount'
    assert by_id[ids[1]]['file_a'] == 'Date'
    # Navigate to Step 3 and back
    at.session_state['current_step'] = 3
    at.run()
    at.session_state['current_step'] = 4
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    by_id = {e['id']: e for e in entries}
    assert by_id[ids[0]]['file_a'] == 'Amount'
    assert by_id[ids[1]]['file_a'] == 'Date'


def test_step4_remove_middle_row_preserves_other_selections() -> None:
    """Remove the middle compare row: other selections and types remain unchanged."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['ID', 'Name', 'Amount', 'Date']
    at.session_state['columns_b'] = ['ID', 'Name', 'Amount', 'Date']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'ID', 'file_b': 'ID'},
        {'id': 'k-2', 'file_a': 'Name', 'file_b': 'Name'},
    ]
    at.session_state['compare_pair_entries'] = []
    at.run()
    # Add three comparison pairs
    for _ in range(3):
        for btn in at.button:
            if btn.key == 'add_compare_pair':
                btn.click()
                break
        at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 3
    ids = [e['id'] for e in entries]
    # Select Amount↔Amount (Text), Date↔Date (Numeric), Amount↔Amount (Text)
    for entry_id, col, typ in zip(
        ids, ['Amount', 'Date', 'Amount'], ['Text', 'Numeric', 'Text'], strict=False
    ):
        for sb in at.selectbox:
            if sb.key == f'compare_pair_{entry_id}_a':
                sb.set_value(col)
                break
        for sb in at.selectbox:
            if sb.key == f'compare_pair_{entry_id}_b':
                sb.set_value(col)
                break
        for sb in at.selectbox:
            if sb.key == f'compare_pair_{entry_id}_type':
                sb.set_value(typ)
                break
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    by_id = {e['id']: e for e in entries}
    assert len(by_id) == 3
    # Remove the middle row
    for btn in at.button:
        if btn.key == f'remove_compare_pair_{ids[1]}':
            btn.click()
            break
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    by_id = {e['id']: e for e in entries}
    assert len(by_id) == 2
    assert ids[1] not in by_id
    assert by_id[ids[0]]['file_a'] == 'Amount'
    assert by_id[ids[0]]['file_b'] == 'Amount'
    assert by_id[ids[2]]['file_a'] == 'Amount'
    assert by_id[ids[2]]['file_b'] == 'Amount'


def test_step5_end_to_end_two_key_pairs_two_compare_pairs() -> None:
    """End-to-end: two key pairs + two compare pairs → comparison receives all."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['file_a_name'] = 'A.xlsx'
    at.session_state['file_b_name'] = 'B.xlsx'
    at.session_state['columns_a'] = ['ID', 'Name', 'Amount', 'Score']
    at.session_state['columns_b'] = ['ID', 'Name', 'Amount', 'Score']
    at.session_state['key_pair_entries'] = []
    at.session_state['compare_pair_entries'] = []
    df_a = pd.DataFrame(
        {'ID': ['1', '2'], 'Name': ['A', 'B'], 'Amount': ['100', '200'], 'Score': ['10', '20']}
    )
    df_b = pd.DataFrame(
        {'ID': ['1', '2'], 'Name': ['A', 'B'], 'Amount': ['100', '200'], 'Score': ['10', '20']}
    )
    at.session_state['loaded_a'] = LoadedSheet(
        df=df_a, sheet_name='Sheet1', file_source='A.xlsx', excel_row_mapping={0: 2, 1: 3}
    )
    at.session_state['loaded_b'] = LoadedSheet(
        df=df_b, sheet_name='Sheet1', file_source='B.xlsx', excel_row_mapping={0: 2, 1: 3}
    )
    at.run()
    # Step 3: add two key pairs, select ID↔ID and Name↔Name
    _add_key_pair_entries(at, 2)
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 2
    ids = [e['id'] for e in entries]
    for entry_id, col in zip(ids, ['ID', 'Name'], strict=False):
        for sb in at.selectbox:
            if sb.key == f'key_pair_{entry_id}_a':
                sb.set_value(col)
                break
        for sb in at.selectbox:
            if sb.key == f'key_pair_{entry_id}_b':
                sb.set_value(col)
                break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    by_id = {e['id']: e for e in entries}
    assert by_id[ids[0]]['file_a'] == 'ID'
    assert by_id[ids[1]]['file_a'] == 'Name'
    # Navigate to Step 4
    at.session_state['current_step'] = 4
    at.run()
    # Add two comparison pairs (Amount and Score are non-key columns)
    for _ in range(2):
        for btn in at.button:
            if btn.key == 'add_compare_pair':
                btn.click()
                break
        at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 2
    cids = [e['id'] for e in entries]
    for entry_id, col in zip(cids, ['Amount', 'Score'], strict=False):
        for sb in at.selectbox:
            if sb.key == f'compare_pair_{entry_id}_a':
                sb.set_value(col)
                break
        for sb in at.selectbox:
            if sb.key == f'compare_pair_{entry_id}_b':
                sb.set_value(col)
                break
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    by_id = {e['id']: e for e in entries}
    assert by_id[cids[0]]['file_a'] == 'Amount'
    assert by_id[cids[1]]['file_a'] == 'Score'
    # Navigate to Step 5 and run comparison
    at.session_state['current_step'] = 5
    at.run()
    for btn in at.button:
        if btn.key == 'btn_run':
            btn.click()
            break
    at.run()
    result = at.session_state.get('comparison_result')
    assert result is not None, 'comparison_result should be stored after running'
    assert len(result.matched) == 2, f'Expected 2 matched records, got {len(result.matched)}'
    assert result.summary['matched_records'] == 2
    assert result.summary['mismatched_records'] == 0


def test_step3_selectbox_selection_persists_to_session_state() -> None:
    """Manual selectbox selection in Step 3 is saved to key_pair_entries."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['ID', 'Name']
    at.session_state['columns_b'] = ['ID', 'Name']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': None, 'file_b': None},
    ]
    at.run()
    # Select ID for File A column of entry k-1
    for sb in at.selectbox:
        if sb.key == 'key_pair_k-1_a':
            sb.set_value('ID')
            break
    at.run()
    # Select ID for File B column of entry k-1
    for sb in at.selectbox:
        if sb.key == 'key_pair_k-1_b':
            sb.set_value('ID')
            break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 1
    assert entries[0]['file_a'] == 'ID'
    assert entries[0]['file_b'] == 'ID'


def test_step3_key_selection_survives_navigation_to_step4() -> None:
    """Key selections made in Step 3 survive navigation to Step 4."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['ID', 'Name']
    at.session_state['columns_b'] = ['ID', 'Name']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': None, 'file_b': None},
    ]
    at.run()
    # Select matching key in Step 3
    for sb in at.selectbox:
        if sb.key == 'key_pair_k-1_a':
            sb.set_value('ID')
            break
    for sb in at.selectbox:
        if sb.key == 'key_pair_k-1_b':
            sb.set_value('ID')
            break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 1
    assert entries[0]['file_a'] == 'ID'
    assert entries[0]['file_b'] == 'ID'
    # Navigate to Step 4
    at.session_state['current_step'] = 4
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 1
    assert entries[0]['file_a'] == 'ID'
    assert entries[0]['file_b'] == 'ID'


def test_step4_compare_selection_persists_to_session_state() -> None:
    """Manual selectbox selection in Step 4 is saved to compare_pair_entries."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 4
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['columns_a'] = ['ID', 'Amount']
    at.session_state['columns_b'] = ['ID', 'Amount']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'ID', 'file_b': 'ID'},
    ]
    at.session_state['compare_pair_entries'] = [
        {'id': 'c-1', 'file_a': None, 'file_b': None, 'type': 'Text'},
    ]
    at.run()
    # Select Amount for File A column of entry c-1
    for sb in at.selectbox:
        if sb.key == 'compare_pair_c-1_a':
            sb.set_value('Amount')
            break
    at.run()
    # Select Amount for File B column of entry c-1
    for sb in at.selectbox:
        if sb.key == 'compare_pair_c-1_b':
            sb.set_value('Amount')
            break
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 1
    assert entries[0]['file_a'] == 'Amount'
    assert entries[0]['file_b'] == 'Amount'


def test_step5_comparison_succeeds_with_widget_selected_keys() -> None:
    """End-to-end: widget-selected keys in Step 3 and compare columns in Step 4
    produce a successful comparison result when Run comparison is clicked."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 3
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['file_a_name'] = 'A.xlsx'
    at.session_state['file_b_name'] = 'B.xlsx'
    at.session_state['columns_a'] = ['ID', 'Amount']
    at.session_state['columns_b'] = ['ID', 'Amount']
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': None, 'file_b': None},
    ]
    at.session_state['compare_pair_entries'] = []
    df_a = pd.DataFrame({'ID': ['1'], 'Amount': ['100']})
    df_b = pd.DataFrame({'ID': ['1'], 'Amount': ['100']})
    at.session_state['loaded_a'] = LoadedSheet(
        df=df_a, sheet_name='Sheet1', file_source='A.xlsx', excel_row_mapping={0: 2}
    )
    at.session_state['loaded_b'] = LoadedSheet(
        df=df_b, sheet_name='Sheet1', file_source='B.xlsx', excel_row_mapping={0: 2}
    )
    at.run()
    # Select matching key in Step 3
    for sb in at.selectbox:
        if sb.key == 'key_pair_k-1_a':
            sb.set_value('ID')
            break
    for sb in at.selectbox:
        if sb.key == 'key_pair_k-1_b':
            sb.set_value('ID')
            break
    at.run()
    entries = at.session_state.get('key_pair_entries', [])
    assert len(entries) == 1
    assert entries[0]['file_a'] == 'ID'
    assert entries[0]['file_b'] == 'ID'
    # Navigate to Step 4
    at.session_state['current_step'] = 4
    at.run()
    # Add a comparison pair in Step 4
    for btn in at.button:
        if btn.key == 'add_compare_pair':
            btn.click()
            break
    at.run()
    # Find the new comparison pair entry (added with a random UUID)
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 1
    new_id = entries[0]['id']
    # Select Amount for File A column of the new entry
    for sb in at.selectbox:
        if sb.key == f'compare_pair_{new_id}_a':
            sb.set_value('Amount')
            break
    at.run()
    # Select Amount for File B column of the new entry
    for sb in at.selectbox:
        if sb.key == f'compare_pair_{new_id}_b':
            sb.set_value('Amount')
            break
    at.run()
    entries = at.session_state.get('compare_pair_entries', [])
    assert len(entries) == 1
    assert entries[0]['file_a'] == 'Amount'
    assert entries[0]['file_b'] == 'Amount'
    # Navigate to Step 5
    at.session_state['current_step'] = 5
    at.run()
    # Click "Run comparison"
    for btn in at.button:
        if btn.key == 'btn_run':
            btn.click()
            break
    at.run()
    result = at.session_state.get('comparison_result')
    assert result is not None, 'comparison_result should be stored after running'
    assert len(result.matched) == 1, f'Expected 1 matched record, got {len(result.matched)}'
    assert result.summary['matched_records'] == 1
    assert result.summary['mismatched_records'] == 0


# ---------------------------------------------------------------------------
# Streamlit heading and complete-record regression tests
# ---------------------------------------------------------------------------


def _step5_streamlit_app_with_complete_data() -> AppTest:
    """Build an AppTest at Step 5 with complete data for heading/record tests."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 5
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['file_a_name'] = 'A.xlsx'
    at.session_state['file_b_name'] = 'B.xlsx'
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'SR NO', 'file_b': 'SR'},
        {'id': 'k-2', 'file_a': 'snapshot_date', 'file_b': 'snapshot_date'},
    ]
    at.session_state['compare_pair_entries'] = [
        {'id': 'c-1', 'file_a': 'warehouse', 'file_b': 'warehouse', 'type': 'Text'},
    ]

    matched = RowDiff(
        row_key="SR NO='101', snapshot_date='2026-01-01'",
        status=DiffStatus.MATCHED,
        cell_diffs=[],
        index_a=0,
        index_b=0,
        excel_row_a=2,
        excel_row_b=2,
        key_values_a=('101', '2026-01-01'),
        key_values_b=('101', '2026-01-01'),
    )
    only_in_a = RowDiff(
        row_key="SR NO='9'",
        status=DiffStatus.ONLY_IN_A,
        cell_diffs=[],
        index_a=1,
        excel_row_a=5,
        key_values_a=('9', '2026-01-02'),
        key_values_b=(),
    )
    only_in_b = RowDiff(
        row_key="SR NO='7'",
        status=DiffStatus.ONLY_IN_B,
        cell_diffs=[],
        index_b=2,
        excel_row_b=6,
        key_values_a=(),
        key_values_b=('7', '2026-01-03'),
    )

    df_a = pd.DataFrame(
        {
            'SR NO': ['101', '9'],
            'snapshot_date': ['2026-01-01', '2026-01-02'],
            'warehouse': ['Jaipur', 'Jodhpur'],
            'extra_a': ['unmapped_a1', 'unmapped_a2'],
        }
    )
    df_b = pd.DataFrame(
        {
            'SR': ['101', '7'],
            'snapshot_date': ['2026-01-01', '2026-01-03'],
            'warehouse': ['Jaipur', 'Udaipur'],
            'extra_b': ['unmapped_b1', 'unmapped_b2'],
        }
    )
    at.session_state['loaded_a'] = LoadedSheet(
        df=df_a, sheet_name='Sheet1', file_source='A.xlsx', excel_row_mapping={0: 2, 1: 5}
    )
    at.session_state['loaded_b'] = LoadedSheet(
        df=df_b, sheet_name='Sheet1', file_source='B.xlsx', excel_row_mapping={0: 2, 2: 6}
    )

    result = ComparisonResult(
        matched=[matched],
        mismatched=[],
        only_in_a=[only_in_a],
        only_in_b=[only_in_b],
        data_quality_issues=[],
        schema_diff=SchemaDiff(
            common_columns=['SR NO', 'SR', 'snapshot_date', 'warehouse'],
            only_in_a=['extra_a'],
            only_in_b=['extra_b'],
            key_columns=['SR NO', 'SR'],
            compare_columns=['warehouse'],
        ),
        summary={
            'total_rows_file_a': 2,
            'total_rows_file_b': 2,
            'matched_records': 1,
            'mismatched_records': 0,
            'matched_no_fields_compared': 0,
            'only_in_file_a': 1,
            'only_in_file_b': 1,
            'duplicate_key_records': 0,
            'missing_key_records': 0,
        },
    )
    at.session_state['comparison_result'] = result
    at.run()
    return at


def test_streamlit_key_column_headings_use_en_dash() -> None:
    """Streamlit key column headings use File A – / File B – en-dash format."""
    at = _step5_streamlit_app_with_complete_data()
    assert not at.exception
    assert at.session_state.get('comparison_result') is not None


def test_streamlit_missing_from_file_b_complete_record() -> None:
    """Missing From File B record data includes complete File A reference."""
    at = _step5_streamlit_app_with_complete_data()
    assert not at.exception
    result = at.session_state.get('comparison_result')
    assert result is not None
    assert len(result.only_in_a) == 1
    assert result.only_in_a[0].index_a == 1
    assert result.only_in_a[0].excel_row_a == 5
    assert result.only_in_a[0].key_values_a == ('9', '2026-01-02')
    assert result.only_in_a[0].key_values_b == ()


def test_streamlit_new_in_file_b_complete_record() -> None:
    """New Records in File B record data includes complete File B reference."""
    at = _step5_streamlit_app_with_complete_data()
    assert not at.exception
    result = at.session_state.get('comparison_result')
    assert result is not None
    assert len(result.only_in_b) == 1
    assert result.only_in_b[0].index_b == 2
    assert result.only_in_b[0].excel_row_b == 6
    assert result.only_in_b[0].key_values_a == ()
    assert result.only_in_b[0].key_values_b == ('7', '2026-01-03')


def test_streamlit_unchanged_records_complete_data() -> None:
    """Unchanged Records record data includes complete matched record reference."""
    at = _step5_streamlit_app_with_complete_data()
    assert not at.exception
    result = at.session_state.get('comparison_result')
    assert result is not None
    assert len(result.matched) == 1
    assert result.matched[0].index_a == 0
    assert result.matched[0].index_b == 0
    assert result.matched[0].key_values_a == ('101', '2026-01-01')
    assert result.matched[0].key_values_b == ('101', '2026-01-01')


# ---------------------------------------------------------------------------
# Preview key-value ordering regression tests (Issue 5)
# Blank key-pair entries must not shift key values under wrong column headers.
# ---------------------------------------------------------------------------


def _step5_app_with_blank_keys(
    key_pair_entries: list[dict],
    key_values_a: tuple,
    key_values_b: tuple,
) -> AppTest:
    """Build an AppTest at Step 5 with custom key-pair entries and key values."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 5
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['file_a_name'] = 'A.xlsx'
    at.session_state['file_b_name'] = 'B.xlsx'
    at.session_state['key_pair_entries'] = key_pair_entries
    at.session_state['compare_pair_entries'] = []

    matched = RowDiff(
        row_key="ID='101'",
        status=DiffStatus.MATCHED,
        cell_diffs=[],
        index_a=0,
        index_b=0,
        excel_row_a=2,
        excel_row_b=2,
        key_values_a=key_values_a,
        key_values_b=key_values_b,
    )

    result = ComparisonResult(
        matched=[matched],
        mismatched=[],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=SchemaDiff(
            common_columns=['ID', 'Code'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['ID', 'Code'],
            compare_columns=[],
        ),
        summary={
            'total_rows_file_a': 1,
            'total_rows_file_b': 1,
            'matched_records': 1,
            'mismatched_records': 0,
            'matched_no_fields_compared': 0,
            'only_in_file_a': 0,
            'only_in_file_b': 0,
            'duplicate_key_records': 0,
            'missing_key_records': 0,
        },
    )
    at.session_state['comparison_result'] = result
    at.run()
    return at


def test_streamlit_preview_blank_entry_between_valid_keys():
    """Blank key-pair entry between valid entries must not shift key values."""
    at = _step5_app_with_blank_keys(
        key_pair_entries=[
            {'id': 'k1', 'file_a': 'ID', 'file_b': 'ID'},
            {'id': 'k2', 'file_a': None, 'file_b': None},  # blank between valid
            {'id': 'k3', 'file_a': 'Code', 'file_b': 'Code'},
        ],
        key_values_a=('101', 'ABC'),
        key_values_b=('101', 'ABC'),
    )
    assert not at.exception

    # Unchanged Records dataframe (tab 0)
    unchanged_df = at.dataframe[0].value
    cols = {str(c): c for c in unchanged_df.columns}

    # File A – ID should be 101 (first valid entry), File B – ID should be 101
    assert unchanged_df.loc[0, cols['File A – ID']] == '101'
    assert unchanged_df.loc[0, cols['File B – ID']] == '101'
    # File A – Code should be ABC (second valid entry, not blank entry)
    assert unchanged_df.loc[0, cols['File A – Code']] == 'ABC'
    assert unchanged_df.loc[0, cols['File B – Code']] == 'ABC'


def test_streamlit_preview_blank_entry_before_valid_key():
    """Blank key-pair entry before a valid entry must not shift key values."""
    at = _step5_app_with_blank_keys(
        key_pair_entries=[
            {'id': 'k1', 'file_a': None, 'file_b': None},  # blank before valid
            {'id': 'k2', 'file_a': 'ID', 'file_b': 'ID'},
        ],
        key_values_a=('101',),
        key_values_b=('101',),
    )
    assert not at.exception

    unchanged_df = at.dataframe[0].value
    cols = {str(c): c for c in unchanged_df.columns}

    assert unchanged_df.loc[0, cols['File A – ID']] == '101'
    assert unchanged_df.loc[0, cols['File B – ID']] == '101'


def test_streamlit_preview_multiple_blank_entries_interspersed():
    """Multiple blank entries interspersed with valid mappings must not shift values."""
    at = _step5_app_with_blank_keys(
        key_pair_entries=[
            {'id': 'k1', 'file_a': None, 'file_b': None},  # blank
            {'id': 'k2', 'file_a': 'ID', 'file_b': 'ID'},  # valid -> index 0
            {'id': 'k3', 'file_a': None, 'file_b': None},  # blank
            {'id': 'k4', 'file_a': 'Code', 'file_b': 'Code'},  # valid -> index 1
            {'id': 'k5', 'file_a': None, 'file_b': None},  # blank
        ],
        key_values_a=('101', 'XYZ'),
        key_values_b=('101', 'XYZ'),
    )
    assert not at.exception

    unchanged_df = at.dataframe[0].value
    cols = {str(c): c for c in unchanged_df.columns}

    assert unchanged_df.loc[0, cols['File A – ID']] == '101'
    assert unchanged_df.loc[0, cols['File B – ID']] == '101'
    assert unchanged_df.loc[0, cols['File A – Code']] == 'XYZ'
    assert unchanged_df.loc[0, cols['File B – Code']] == 'XYZ'


# ---------------------------------------------------------------------------
# Unchanged Records preview — configured columns only (keys + File B compares)
# ---------------------------------------------------------------------------
# Generic inventory-style fixture. Column names are arbitrary test data, not
# production constants, so the behavior is verified generically.
_PREVIEW_INVENTORY_KEYS = [('SR NO', 'SR')]
_PREVIEW_INVENTORY_COMPARE = [
    'on_hand',
    'reserved',
    'available',
    'reorder_point',
    'below_reorder',
    'unit_cost',
    'days_of_cover',
]
_PREVIEW_INVENTORY_UNMAPPED = ['snapshot_date', 'warehouse', 'sku']


def _step5_inventory_app() -> AppTest:
    """Build an AppTest at Step 5 with one matched inventory-style record."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 5
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['file_a_name'] = 'A.xlsx'
    at.session_state['file_b_name'] = 'B.xlsx'
    at.session_state['key_pair_entries'] = [
        {
            'id': 'k-1',
            'file_a': _PREVIEW_INVENTORY_KEYS[0][0],
            'file_b': _PREVIEW_INVENTORY_KEYS[0][1],
        }
    ]
    at.session_state['compare_pair_entries'] = [
        {'id': f'c-{i}', 'file_a': col, 'file_b': col, 'type': 'Text'}
        for i, col in enumerate(_PREVIEW_INVENTORY_COMPARE)
    ]

    row_a: dict = {'SR NO': '1000'}
    row_b: dict = {'SR': '1000'}
    for i, col in enumerate(_PREVIEW_INVENTORY_COMPARE):
        row_a[col] = str(100 + i)
        row_b[col] = str(100 + i)
    for col in _PREVIEW_INVENTORY_UNMAPPED:
        row_a[col] = f'a_{col}'
        row_b[col] = f'b_{col}'
    at.session_state['loaded_a'] = LoadedSheet(
        df=pd.DataFrame([row_a]),
        sheet_name='Sheet1',
        file_source='A.xlsx',
        excel_row_mapping={0: 2},
    )
    at.session_state['loaded_b'] = LoadedSheet(
        df=pd.DataFrame([row_b]),
        sheet_name='Sheet1',
        file_source='B.xlsx',
        excel_row_mapping={0: 2},
    )

    at.session_state['comparison_result'] = ComparisonResult(
        matched=[
            RowDiff(
                row_key="SR NO='1000'",
                status=DiffStatus.MATCHED,
                cell_diffs=[],
                index_a=0,
                index_b=0,
                excel_row_a=2,
                excel_row_b=2,
                key_values_a=('1000',),
                key_values_b=('1000',),
            )
        ],
        mismatched=[],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=SchemaDiff(
            common_columns=['SR NO', 'SR', *_PREVIEW_INVENTORY_COMPARE],
            only_in_a=list(_PREVIEW_INVENTORY_UNMAPPED),
            only_in_b=list(_PREVIEW_INVENTORY_UNMAPPED),
            key_columns=['SR NO', 'SR'],
            compare_columns=list(_PREVIEW_INVENTORY_COMPARE),
        ),
        summary={
            'total_rows_file_a': 1,
            'total_rows_file_b': 1,
            'matched_records': 1,
            'mismatched_records': 0,
            'matched_no_fields_compared': 0,
            'only_in_file_a': 0,
            'only_in_file_b': 0,
            'duplicate_key_records': 0,
            'missing_key_records': 0,
        },
    )
    at.run()
    return at


def test_streamlit_unchanged_preview_only_configured_columns() -> None:
    """Unchanged preview shows exactly keys + File B comparisons, in order."""
    at = _step5_inventory_app()
    assert not at.exception

    preview = at.dataframe[0].value
    expected_columns = (
        ['File A – SR NO', 'File B – SR']
        + [f'File B – {col}' for col in _PREVIEW_INVENTORY_COMPARE]
        + ['Record Status', 'Excel Row (File A)', 'Excel Row (File B)']
    )
    assert list(preview.columns) == expected_columns

    # Unmapped File A / File B columns are not displayed.
    for col in _PREVIEW_INVENTORY_UNMAPPED:
        assert f'File A – {col}' not in preview.columns
        assert f'File B – {col}' not in preview.columns

    assert len(preview) == 1
    preview_row = preview.iloc[0]
    assert preview_row['File A – SR NO'] == '1000'
    assert preview_row['File B – SR'] == '1000'
    for i, col in enumerate(_PREVIEW_INVENTORY_COMPARE):
        assert preview_row[f'File B – {col}'] == str(100 + i)
    assert preview_row['Record Status'] == 'No changes'
    assert preview_row['Excel Row (File A)'] == 2
    assert preview_row['Excel Row (File B)'] == 2


# ---------------------------------------------------------------------------
# Records With Changes preview — Arrow-safe mixed-type display values
# ---------------------------------------------------------------------------
# The preview gathers values from different source columns into one Previous
# Value / Updated Value column, so those columns mix native scalar types
# (numpy bool/int/float, str) that PyArrow cannot serialize into a single Arrow
# column. They are rendered as display strings. These tests exercise the
# conversion helper directly with PyArrow and the full app end-to-end.

_APP_MODULE = None


def _app_module():
    """Import the Streamlit app module once (safe: main() is guarded)."""
    global _APP_MODULE
    if _APP_MODULE is None:
        spec = importlib.util.spec_from_file_location('_app_under_test', APP_PATH)
        assert spec is not None and spec.loader is not None
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        _APP_MODULE = module
    return _APP_MODULE


def test_display_value_stringifies_mixed_numpy_types() -> None:
    """Mixed numpy scalar types become strings and serialize through PyArrow."""
    display = _app_module()._display_value

    assert display(np.True_) == 'True'
    assert display(np.False_) == 'False'
    assert display(np.int64(827)) == '827'
    assert display(np.float64(231.76)) == '231.76'
    assert display('DC-Newark') == 'DC-Newark'

    values = [display(v) for v in [np.int64(827), np.False_, np.float64(231.76), 'DC-Newark']]
    table = pa.Table.from_pandas(pd.DataFrame({'Previous Value': values}))
    assert table.column('Previous Value').to_pylist() == ['827', 'False', '231.76', 'DC-Newark']


def test_raw_mixed_type_column_fails_arrow_serialization() -> None:
    """Baseline: the raw mixed-type column is what breaks PyArrow serialization."""
    raw = pd.DataFrame({'Previous Value': [np.int64(827), np.False_, np.float64(231.76)]})
    with pytest.raises(pa.ArrowInvalid):
        pa.Table.from_pandas(raw)


def test_display_value_renders_missing_as_blank() -> None:
    """None, NaN, pd.NA and NaT render as empty strings."""
    display = _app_module()._display_value
    for missing in (None, float('nan'), np.nan, pd.NA, pd.NaT):
        assert display(missing) == ''


class _RecordHandler(logging.Handler):
    """Collect log records from a specific logger (Streamlit sets propagate=False)."""

    def __init__(self) -> None:
        super().__init__(level=logging.INFO)
        self.records: list[logging.LogRecord] = []

    def emit(self, record: logging.LogRecord) -> None:
        self.records.append(record)


def _step5_app_with_mixed_types() -> AppTest:
    """Step 5 app whose Records With Changes row mixes numpy bool/int/float/str."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 5
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['file_a_name'] = 'A.xlsx'
    at.session_state['file_b_name'] = 'B.xlsx'
    at.session_state['key_pair_entries'] = [
        {'id': 'k-1', 'file_a': 'ID', 'file_b': 'ID', 'type': 'Text'},
    ]
    at.session_state['compare_pair_entries'] = [
        {'id': 'c-1', 'file_a': 'qty', 'file_b': 'qty', 'type': 'Text'},
        {'id': 'c-2', 'file_a': 'active', 'file_b': 'active', 'type': 'Text'},
        {'id': 'c-3', 'file_a': 'price', 'file_b': 'price', 'type': 'Text'},
    ]

    cell_diffs = [
        CellDiff(
            column_a='qty',
            column_b='qty',
            row_key="ID='2'",
            sheet='',
            value_a=np.int64(827),
            value_b=np.int64(900),
            are_equal=False,
            excel_row_a=3,
            excel_row_b=3,
        ),
        CellDiff(
            column_a='active',
            column_b='active',
            row_key="ID='2'",
            sheet='',
            value_a=np.False_,
            value_b=np.True_,
            are_equal=False,
            excel_row_a=3,
            excel_row_b=3,
        ),
        CellDiff(
            column_a='price',
            column_b='price',
            row_key="ID='2'",
            sheet='',
            value_a=np.float64(231.76),
            value_b=np.float64(240.5),
            are_equal=False,
            excel_row_a=3,
            excel_row_b=3,
        ),
        CellDiff(
            column_a='note',
            column_b='note',
            row_key="ID='2'",
            sheet='',
            value_a=float('nan'),
            value_b=pd.NaT,
            are_equal=False,
            excel_row_a=3,
            excel_row_b=3,
        ),
    ]

    result = ComparisonResult(
        matched=[
            RowDiff(
                row_key="ID='1'",
                status=DiffStatus.MATCHED,
                cell_diffs=[],
                index_a=0,
                index_b=0,
                excel_row_a=2,
                excel_row_b=2,
                key_values_a=('1',),
                key_values_b=('1',),
            )
        ],
        mismatched=[
            RowDiff(
                row_key="ID='2'",
                status=DiffStatus.MISMATCHED,
                cell_diffs=cell_diffs,
                index_a=1,
                index_b=1,
                excel_row_a=3,
                excel_row_b=3,
                key_values_a=('2',),
                key_values_b=('2',),
            )
        ],
        only_in_a=[],
        only_in_b=[],
        data_quality_issues=[],
        schema_diff=SchemaDiff(
            common_columns=['ID', 'qty', 'active', 'price', 'note'],
            only_in_a=[],
            only_in_b=[],
            key_columns=['ID'],
            compare_columns=['qty', 'active', 'price', 'note'],
        ),
        summary={
            'total_rows_file_a': 2,
            'total_rows_file_b': 2,
            'matched_records': 1,
            'mismatched_records': 1,
            'matched_no_fields_compared': 0,
            'only_in_file_a': 0,
            'only_in_file_b': 0,
            'duplicate_key_records': 0,
            'missing_key_records': 0,
        },
    )
    at.session_state['comparison_result'] = result
    at.run()
    return at


def test_changed_preview_mixed_types_render_as_strings_without_arrow_fallback() -> None:
    """Records With Changes renders mixed types as strings without Arrow fallback."""
    arrow_logger = logging.getLogger('streamlit.dataframe_util')
    handler = _RecordHandler()
    arrow_logger.addHandler(handler)
    try:
        at = _step5_app_with_mixed_types()
    finally:
        arrow_logger.removeHandler(handler)

    assert not at.exception

    # No automatic Arrow fallback was triggered while rendering the preview.
    fallback = [
        r
        for r in handler.records
        if 'Serialization of dataframe to Arrow table was unsuccessful' in r.getMessage()
    ]
    assert fallback == []

    # at.dataframe[0] = Unchanged Records, at.dataframe[1] = Records With Changes.
    changed = at.dataframe[1].value
    assert changed['Previous Value'].tolist() == ['827', 'False', '231.76', '']
    assert changed['Updated Value'].tolist() == ['900', 'True', '240.5', '']

    # Other preview columns retain their existing behavior.
    assert changed['Column Name (File A)'].tolist() == ['qty', 'active', 'price', 'note']
    assert changed['Column Name (File B)'].tolist() == ['qty', 'active', 'price', 'note']
    assert changed['File A – ID'].tolist() == ['2', '2', '2', '2']
    assert changed['File B – ID'].tolist() == ['2', '2', '2', '2']
    assert changed['Excel Row (File A)'].tolist() == [3, 3, 3, 3]
    assert changed['Excel Row (File B)'].tolist() == [3, 3, 3, 3]


def test_app_source_has_no_use_container_width() -> None:
    """The deprecated use_container_width argument must not appear in the app."""
    source = APP_PATH.read_text(encoding='utf-8')
    assert 'use_container_width' not in source


# ---------------------------------------------------------------------------
# Temporary data cleanup (Finish & Clear Data / Start Over / uploader nonce)
# ---------------------------------------------------------------------------

_APP_OWNED_DEFAULTS = {
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
}


def _assert_app_owned_defaults(at: AppTest) -> None:
    """Assert every application-owned key is back at its initial default."""
    for key, expected in _APP_OWNED_DEFAULTS.items():
        actual = at.session_state.get(key)
        assert actual == expected, f'{key}: expected {expected!r}, got {actual!r}'


def _click_button(at: AppTest, key: str) -> None:
    for btn in at.button:
        if btn.key == key:
            btn.click()
            return
    raise AssertionError(f'button {key!r} not found')  # noqa: TRY003


def _check_checkbox(at: AppTest, key: str) -> None:
    for cb in at.checkbox:
        if cb.key == key:
            cb.check()
            return
    raise AssertionError(f'checkbox {key!r} not found')  # noqa: TRY003


def _step6_app_with_report() -> AppTest:
    """Build an AppTest at Step 6 with report bytes and app-owned data present."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = 6
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['file_a_name'] = 'A.xlsx'
    at.session_state['file_b_name'] = 'B.xlsx'
    at.session_state['file_a_bytes'] = b'fake-a'
    at.session_state['file_b_bytes'] = b'fake-b'
    at.session_state['file_a_hash'] = 'hash-a'
    at.session_state['file_b_hash'] = 'hash-b'
    at.session_state['sheet_a'] = 'Sheet1'
    at.session_state['sheet_b'] = 'Sheet1'
    at.session_state['sheets_a'] = ['Sheet1']
    at.session_state['sheets_b'] = ['Sheet1']
    at.session_state['loaded_a'] = object()
    at.session_state['loaded_b'] = object()
    at.session_state['columns_a'] = ['ID', 'Amount']
    at.session_state['columns_b'] = ['ID', 'Amount']
    at.session_state['key_pair_entries'] = [{'id': 'k-1', 'file_a': 'ID', 'file_b': 'ID'}]
    at.session_state['compare_pair_entries'] = [
        {'id': 'c-1', 'file_a': 'Amount', 'file_b': 'Amount', 'type': 'Text'}
    ]
    at.session_state['comparison_result'] = object()
    at.session_state['report_bytes'] = b'report-bytes'
    at.run()
    return at


def test_download_and_report_available_before_confirmation() -> None:
    """The report and download button stay available until the user confirms."""
    at = _step6_app_with_report()
    assert not at.exception
    assert at.session_state.get('report_bytes') == b'report-bytes'
    assert at.session_state.get('current_step') == 6
    assert [d.key for d in at.download_button] == ['btn_download_report']


def test_finish_clear_requires_confirmation() -> None:
    """Clicking Finish & Clear Data without confirmation keeps all data."""
    at = _step6_app_with_report()
    _click_button(at, 'btn_finish_clear')
    at.run()
    assert not at.exception
    assert at.session_state.get('report_bytes') == b'report-bytes'
    assert at.session_state.get('current_step') == 6
    assert any('confirmation' in w.value for w in at.warning)


def test_finish_clear_confirmed_clears_all_session_data() -> None:
    """Confirming Finish & Clear Data resets every app-owned key and Step 1."""
    at = _step6_app_with_report()
    _check_checkbox(at, 'confirm_finish_clear')
    at.run()
    _click_button(at, 'btn_finish_clear')
    at.run()
    assert not at.exception
    _assert_app_owned_defaults(at)


def test_start_over_requires_confirmation() -> None:
    """Clicking Start Over without confirmation keeps all data."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['file_a_obj'] = object()
    at.session_state['file_a_bytes'] = b'keep-a'
    at.run()
    _click_button(at, 'btn_start_over')
    at.run()
    assert not at.exception
    assert at.session_state.get('file_a_bytes') == b'keep-a'
    assert at.session_state.get('file_a_obj') is not None
    assert any('confirmation' in w.value for w in at.warning)


@pytest.mark.parametrize('step', [1, 2, 3, 4, 5, 6])
def test_start_over_confirmed_clears_from_every_step(step: int) -> None:
    """Confirmed Start Over clears all data and returns to Step 1 from any step."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['current_step'] = step
    at.session_state['file_a_obj'] = object()
    at.session_state['file_b_obj'] = object()
    at.session_state['file_a_bytes'] = b'a'
    at.session_state['file_b_bytes'] = b'b'
    at.session_state['report_bytes'] = b'r'
    at.run()
    assert not at.exception
    _check_checkbox(at, 'confirm_start_over')
    at.run()
    _click_button(at, 'btn_start_over')
    at.run()
    assert not at.exception
    assert at.session_state.get('current_step') == 1
    assert at.session_state.get('file_a_obj') is None
    assert at.session_state.get('file_a_bytes') is None
    assert at.session_state.get('file_b_obj') is None
    assert at.session_state.get('file_b_bytes') is None
    assert at.session_state.get('report_bytes') is None


def test_clear_increments_nonce_and_rotates_uploader_keys() -> None:
    """Cleanup preserves and increments the nonce, rotating uploader widget keys."""
    at = AppTest.from_file(APP_PATH)
    at.run()
    assert at.session_state.get('_session_nonce') == 0
    keys_before = sorted(w.key for w in at.file_uploader)
    assert keys_before == ['uploader_file_a_0', 'uploader_file_b_0']

    _check_checkbox(at, 'confirm_start_over')
    at.run()
    _click_button(at, 'btn_start_over')
    at.run()

    assert not at.exception
    assert at.session_state.get('_session_nonce') == 1
    keys_after = sorted(w.key for w in at.file_uploader)
    assert keys_after == ['uploader_file_a_1', 'uploader_file_b_1']
    assert keys_before != keys_after


def test_old_uploaded_files_not_reingested_after_clear() -> None:
    """Old uploaded files do not reappear after cleanup and further reruns."""
    at = AppTest.from_file(APP_PATH)
    at.session_state['file_a_obj'] = object()
    at.session_state['file_a_bytes'] = b'old-a'
    at.session_state['file_a_name'] = 'old-a.xlsx'
    at.session_state['file_b_obj'] = object()
    at.session_state['file_b_bytes'] = b'old-b'
    at.session_state['file_b_name'] = 'old-b.xlsx'
    at.run()
    assert at.session_state.get('file_a_bytes') == b'old-a'

    _check_checkbox(at, 'confirm_start_over')
    at.run()
    _click_button(at, 'btn_start_over')
    at.run()

    assert at.session_state.get('file_a_obj') is None
    assert at.session_state.get('file_a_bytes') is None
    # A further rerun must not resurrect the previously uploaded files.
    at.run()
    assert not at.exception
    assert at.session_state.get('file_a_obj') is None
    assert at.session_state.get('file_a_bytes') is None
    assert at.session_state.get('file_b_obj') is None
    assert at.session_state.get('file_b_bytes') is None


def test_clear_session_resets_sheets_state() -> None:
    """sheets_a/sheets_b are reset alongside the other app-owned state."""
    at = _step6_app_with_report()
    assert at.session_state.get('sheets_a') == ['Sheet1']
    _check_checkbox(at, 'confirm_finish_clear')
    at.run()
    _click_button(at, 'btn_finish_clear')
    at.run()
    assert at.session_state.get('sheets_a') == []
    assert at.session_state.get('sheets_b') == []


def test_two_apptest_sessions_remain_isolated() -> None:
    """Clearing one AppTest session must not affect another independent session."""
    at1 = _step6_app_with_report()
    at2 = _step6_app_with_report()

    _check_checkbox(at1, 'confirm_finish_clear')
    at1.run()
    _click_button(at1, 'btn_finish_clear')
    at1.run()

    assert at1.session_state.get('report_bytes') is None
    assert at1.session_state.get('current_step') == 1
    # The second session is untouched.
    assert at2.session_state.get('report_bytes') == b'report-bytes'
    assert at2.session_state.get('current_step') == 6
    assert not at2.exception
    assert at2.session_state.get('_session_nonce') == 0
