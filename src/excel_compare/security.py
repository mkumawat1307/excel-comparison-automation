"""Security validation for uploaded Excel files.

Guarantees:
- Only .xlsx files are accepted.
- File size is bounded (default 50 MB).
- The workbook is readable before further processing.
- Original uploaded files are never modified.
"""

import shutil
import zipfile
from pathlib import Path
from typing import Union

import openpyxl

from .config import ALLOWED_EXTENSIONS, MAX_DECOMPRESSED_SIZE_BYTES, MAX_FILE_SIZE_BYTES
from .exceptions import SecurityError


def validate_uploaded_file(file_path: Union[str, Path]) -> Path:
    """Validate an uploaded file before processing.

    Returns the resolved Path if validation succeeds.
    Raises SecurityError on any failure.
    """
    path = Path(file_path).resolve()

    if not path.exists():
        raise SecurityError(f'File does not exist: {path}')  # noqa: TRY003

    if path.is_dir():
        raise SecurityError(f'Path is a directory, not a file: {path}')  # noqa: TRY003

    # Extension check (case-insensitive)
    if path.suffix.lower() not in ALLOWED_EXTENSIONS:
        raise SecurityError(  # noqa: TRY003
            f"Unsupported file extension '{path.suffix}'. Allowed: {', '.join(ALLOWED_EXTENSIONS)}"
        )

    # Size check
    size = path.stat().st_size
    if size == 0:
        raise SecurityError(f'File is empty: {path}')  # noqa: TRY003
    if size > MAX_FILE_SIZE_BYTES:
        raise SecurityError(  # noqa: TRY003
            f'File exceeds maximum size of {MAX_FILE_SIZE_BYTES // (1024 * 1024)} MB: {path}'
        )

    # Decompressed size check (ZIP bomb protection)
    try:
        with zipfile.ZipFile(path) as zf:
            decompressed_size = sum(info.file_size for info in zf.infolist())
        if decompressed_size > MAX_DECOMPRESSED_SIZE_BYTES:
            raise SecurityError(  # noqa: TRY003
                f'File decompressed size exceeds maximum of '
                f'{MAX_DECOMPRESSED_SIZE_BYTES // (1024 * 1024)} MB: {path}'
            )
    except zipfile.BadZipFile as exc:
        raise SecurityError(  # noqa: TRY003
            f'File is not a valid ZIP archive: {path} ({exc})'
        ) from exc

    # Readability check — attempt to load the workbook without modifying it
    try:
        openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:  # noqa: BLE001 — any workbook error is a security/load failure
        raise SecurityError(  # noqa: TRY003
            f'File is not a readable Excel workbook: {path} ({exc})'
        ) from exc

    return path


def safe_copy_to_temp(source_path: Union[str, Path], temp_dir: Union[str, Path]) -> Path:
    """Copy a validated file into an isolated temp location.

    The original file is never modified. Returns the temp copy path.
    """
    src = Path(source_path)
    dst = Path(temp_dir) / src.name
    shutil.copy2(src, dst)
    return dst
