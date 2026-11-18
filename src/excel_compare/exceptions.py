"""Custom exceptions for the Excel comparison application."""


class ExcelCompareError(Exception):
    """Base exception for all application errors."""


class LoadError(ExcelCompareError):
    """Raised when an Excel file cannot be loaded or read."""


class ConfigError(ExcelCompareError):
    """Raised when configuration is invalid."""


class SecurityError(ExcelCompareError):
    """Raised when a file fails security validation."""


class DuplicateKeyError(ExcelCompareError):
    """Raised when duplicate keys are detected and policy is 'error' or unsupported."""


class InvalidKeySelectionError(ExcelCompareError):
    """Raised when selected key columns are missing from one or both files."""


class MappingValidationError(ExcelCompareError):
    """Raised when column mapping configuration is invalid."""
