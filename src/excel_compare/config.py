"""Application defaults and constants (no external config files)."""

from typing import Optional

from .models import ColumnPair

DEFAULT_KEYS: list[str] = []
DEFAULT_COMPARE_COLUMNS: Optional[list[str]] = None
DEFAULT_NUMERIC_COLUMNS: Optional[list[str]] = None
DEFAULT_DATE_COLUMNS: Optional[list[str]] = None
DEFAULT_KEY_PAIRS: list[ColumnPair] = []
DEFAULT_COMPARE_PAIRS: list[ColumnPair] = []
DEFAULT_NUMERIC_PAIRS: list[ColumnPair] = []
DEFAULT_DATE_PAIRS: list[ColumnPair] = []
DEFAULT_CASE_SENSITIVE: bool = False
DEFAULT_NORMALIZE_WHITESPACE: bool = True
DEFAULT_NUMERIC_TOLERANCE: float = 1e-9
DEFAULT_DUPLICATE_POLICY: str = 'error'
VALID_DUPLICATE_POLICIES: tuple[str, ...] = ('error',)
DEFAULT_MAX_FILE_SIZE_MB: int = 50
ALLOWED_EXTENSIONS: tuple[str, ...] = ('.xlsx',)
MAX_FILE_SIZE_BYTES: int = 50 * 1024 * 1024

REPORT_SHEETS: tuple[str, ...] = (
    'Comparison Overview',
    'Unchanged Records',
    'Records With Changes',
    'Missing From File B',
    'New Records in File B',
    'Records Needing Attention',
)
