"""Custom exceptions for Excel to PowerPoint Merger."""

from typing import Optional


class ExcelPptxMergerError(Exception):
    """Base exception for Excel to PowerPoint Merger operations."""

    def __init__(self, message: str, error_code: Optional[str] = None) -> None:
        super().__init__(message)
        self.message = message
        self.error_code = error_code


class FileProcessingError(ExcelPptxMergerError):
    """Exception raised for file processing errors."""

    pass


class ExcelProcessingError(FileProcessingError):
    """Exception raised for Excel file processing errors."""

    pass


class PowerPointProcessingError(FileProcessingError):
    """Exception raised for PowerPoint file processing errors."""

    pass


class ConfigurationError(ExcelPptxMergerError):
    """Exception raised for configuration-related errors."""

    pass


class ValidationError(ExcelPptxMergerError):
    """Exception raised for validation errors."""

    pass


class TemplateError(ExcelPptxMergerError):
    """Exception raised for template processing errors."""

    pass


class TempFileError(ExcelPptxMergerError):
    """Exception raised for temporary file management errors."""

    pass


class APIError(ExcelPptxMergerError):
    """Exception raised for API-related errors."""

    pass


class AuthenticationError(APIError):
    """Exception raised for authentication errors."""

    pass


class RateLimitError(APIError):
    """Exception raised for rate limiting errors."""

    pass


class ExternalServiceError(ExcelPptxMergerError):
    """Exception raised for external service errors."""

    pass
