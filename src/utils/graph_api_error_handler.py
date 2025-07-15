"""Error handling utilities for Graph API operations."""

import logging
import time
from functools import wraps
from typing import Any, Callable, Dict, Optional, Type, Union
import requests

from .exceptions import ExcelProcessingError

logger = logging.getLogger(__name__)


class GraphAPIRetryableError(Exception):
    """Exception for retryable Graph API errors."""

    pass


class GraphAPIFatalError(Exception):
    """Exception for non-retryable Graph API errors."""

    pass


class GraphAPIErrorHandler:
    """Handles Graph API errors with retry logic and categorization."""

    # HTTP status codes that are retryable
    RETRYABLE_STATUS_CODES = {
        429,  # Too Many Requests
        500,  # Internal Server Error
        502,  # Bad Gateway
        503,  # Service Unavailable
        504,  # Gateway Timeout
    }

    # Graph API error codes that are retryable
    RETRYABLE_ERROR_CODES = {
        "TooManyRequests",
        "ServiceNotAvailable",
        "Timeout",
        "InternalServerError",
    }

    def __init__(
        self, max_retries: int = 3, base_delay: float = 1.0, max_delay: float = 60.0
    ):
        """Initialize error handler with retry settings."""
        self.max_retries = max_retries
        self.base_delay = base_delay
        self.max_delay = max_delay

    def is_retryable_error(self, error: Exception) -> bool:
        """Determine if an error is retryable."""
        if isinstance(error, requests.exceptions.RequestException):
            if hasattr(error, "response") and error.response is not None:
                status_code = error.response.status_code
                return status_code in self.RETRYABLE_STATUS_CODES

        if isinstance(error, GraphAPIRetryableError):
            return True

        return False

    def get_retry_delay(self, attempt: int, error: Exception = None) -> float:
        """Calculate retry delay with exponential backoff."""
        # Check for Retry-After header in HTTP 429 responses
        if (
            error
            and hasattr(error, "response")
            and error.response is not None
            and error.response.status_code == 429
        ):

            retry_after = error.response.headers.get("Retry-After")
            if retry_after:
                try:
                    return min(float(retry_after), self.max_delay)
                except ValueError:
                    pass

        # Exponential backoff
        delay = self.base_delay * (2**attempt)
        return min(delay, self.max_delay)

    def categorize_error(self, error: Exception) -> Dict[str, Any]:
        """Categorize Graph API error for better handling."""
        error_info = {
            "type": type(error).__name__,
            "message": str(error),
            "retryable": self.is_retryable_error(error),
            "category": "unknown",
        }

        if isinstance(error, requests.exceptions.RequestException):
            error_info["category"] = "http"
            if hasattr(error, "response") and error.response is not None:
                error_info["status_code"] = error.response.status_code
                error_info["response_text"] = error.response.text[:500]  # Limit size

                # Try to parse Graph API error response
                try:
                    error_json = error.response.json()
                    if "error" in error_json:
                        graph_error = error_json["error"]
                        error_info["graph_code"] = graph_error.get("code", "")
                        error_info["graph_message"] = graph_error.get("message", "")

                        # Check if Graph API error code is retryable
                        if graph_error.get("code") in self.RETRYABLE_ERROR_CODES:
                            error_info["retryable"] = True

                except Exception:
                    pass

        elif isinstance(error, (ConnectionError, TimeoutError)):
            error_info["category"] = "network"
            error_info["retryable"] = True

        elif isinstance(error, GraphAPIRetryableError):
            error_info["category"] = "graph_api"
            error_info["retryable"] = True

        elif isinstance(error, GraphAPIFatalError):
            error_info["category"] = "graph_api"
            error_info["retryable"] = False

        return error_info

    def handle_error(
        self, error: Exception, operation: str, context: Dict[str, Any] = None
    ) -> None:
        """Handle and log Graph API error with context."""
        error_info = self.categorize_error(error)
        context = context or {}

        log_message = f"Graph API error in {operation}: {error_info['message']}"
        if context:
            log_message += f" | Context: {context}"

        if error_info["retryable"]:
            logger.warning(f"Retryable {log_message}")
        else:
            logger.error(f"Fatal {log_message}")

        # Log additional details at debug level
        logger.debug(f"Error details: {error_info}")


def with_retry(
    max_retries: int = 3,
    base_delay: float = 1.0,
    error_handler: Optional[GraphAPIErrorHandler] = None,
):
    """Decorator for adding retry logic to Graph API operations."""

    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs) -> Any:
            handler = error_handler or GraphAPIErrorHandler(max_retries, base_delay)
            last_error = None

            for attempt in range(max_retries + 1):
                try:
                    return func(*args, **kwargs)

                except Exception as e:
                    last_error = e

                    if attempt == max_retries:
                        # Final attempt failed
                        handler.handle_error(
                            e,
                            func.__name__,
                            {
                                "attempt": attempt + 1,
                                "max_retries": max_retries,
                                "args": str(args)[:200],
                                "kwargs": str(kwargs)[:200],
                            },
                        )
                        raise ExcelProcessingError(
                            f"Graph API operation failed after {max_retries + 1} attempts: {e}"
                        )

                    if not handler.is_retryable_error(e):
                        # Non-retryable error
                        handler.handle_error(
                            e,
                            func.__name__,
                            {"attempt": attempt + 1, "retryable": False},
                        )
                        raise ExcelProcessingError(
                            f"Graph API operation failed with non-retryable error: {e}"
                        )

                    # Calculate delay and retry
                    delay = handler.get_retry_delay(attempt, e)
                    logger.info(
                        f"Retrying {func.__name__} in {delay:.1f}s (attempt {attempt + 2}/{max_retries + 1})"
                    )
                    time.sleep(delay)

            # Should never reach here, but just in case
            raise ExcelProcessingError(f"Graph API operation failed: {last_error}")

        return wrapper

    return decorator


def safe_graph_operation(
    operation_name: str, error_handler: Optional[GraphAPIErrorHandler] = None
):
    """Context manager for safe Graph API operations with error handling."""

    class GraphOperationContext:
        def __init__(self, name: str, handler: Optional[GraphAPIErrorHandler]):
            self.name = name
            self.handler = handler or GraphAPIErrorHandler()
            self.start_time = None

        def __enter__(self):
            self.start_time = time.time()
            logger.debug(f"Starting Graph API operation: {self.name}")
            return self

        def __exit__(self, exc_type, exc_val, exc_tb):
            duration = time.time() - self.start_time if self.start_time else 0

            if exc_type is None:
                logger.debug(
                    f"Graph API operation completed: {self.name} ({duration:.2f}s)"
                )
            else:
                self.handler.handle_error(
                    exc_val, self.name, {"duration": f"{duration:.2f}s"}
                )

            return False  # Don't suppress exceptions

    return GraphOperationContext(operation_name, error_handler)


def validate_graph_response(response: requests.Response, operation: str) -> None:
    """Validate Graph API response and raise appropriate errors."""
    try:
        response.raise_for_status()
    except requests.exceptions.HTTPError as e:
        if response.status_code in GraphAPIErrorHandler.RETRYABLE_STATUS_CODES:
            raise GraphAPIRetryableError(
                f"{operation} failed with retryable error: {e}"
            )
        else:
            raise GraphAPIFatalError(f"{operation} failed with fatal error: {e}")

    # Additional validation for Graph API specific errors
    try:
        if response.headers.get("content-type", "").startswith("application/json"):
            data = response.json()
            if "error" in data:
                error_code = data["error"].get("code", "")
                error_message = data["error"].get("message", "")

                if error_code in GraphAPIErrorHandler.RETRYABLE_ERROR_CODES:
                    raise GraphAPIRetryableError(f"{operation} failed: {error_message}")
                else:
                    raise GraphAPIFatalError(f"{operation} failed: {error_message}")
    except ValueError:
        # Not JSON response, continue
        pass
