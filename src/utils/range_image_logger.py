"""Specialized logger for range image debugging with enhanced visibility."""

import logging
import sys
from typing import Optional


class RangeImageLogger:
    """Specialized logger for range image operations with enhanced visibility."""

    def __init__(self, name: str = "range_images", level: int = logging.INFO):
        """Initialize the range image logger."""
        self.logger = logging.getLogger(name)
        self.logger.setLevel(level)
        self.debug_enabled = False

        # Remove existing handlers to avoid duplicates
        for handler in self.logger.handlers[:]:
            self.logger.removeHandler(handler)

        # Create console handler with special formatting
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(level)

        # Create formatter with clear visual separation
        formatter = logging.Formatter(
            "\n" + "=" * 80 + "\n"
            "🖼️  RANGE IMAGE DEBUG: %(levelname)s\n"
            "📍 %(name)s - %(funcName)s:%(lineno)d\n"
            "💬 %(message)s\n" + "=" * 80 + "\n"
        )
        console_handler.setFormatter(formatter)

        self.logger.addHandler(console_handler)
        self.logger.propagate = False  # Prevent propagation to root logger

    def enable_debug(self):
        """Enable debug logging for range extraction."""
        self.debug_enabled = True

    def disable_debug(self):
        """Disable debug logging for range extraction."""
        self.debug_enabled = False

    def info(self, message: str, **kwargs) -> None:
        """Log info message with enhanced visibility."""
        if self.debug_enabled:
            self.logger.info(message, **kwargs)

    def warning(self, message: str, **kwargs) -> None:
        """Log warning message with enhanced visibility."""
        self.logger.warning(message, **kwargs)

    def error(self, message: str, **kwargs) -> None:
        """Log error message with enhanced visibility."""
        self.logger.error(message, **kwargs)

    def debug(self, message: str, **kwargs) -> None:
        """Log debug message with enhanced visibility."""
        self.logger.debug(message, **kwargs)

    def critical(self, message: str, **kwargs) -> None:
        """Log critical message with enhanced visibility."""
        self.logger.critical(message, **kwargs)


# Global instance for easy import
range_image_logger = RangeImageLogger()


def setup_range_image_debug_mode(
    enabled: bool = True, level: int = logging.DEBUG
) -> None:
    """Setup range image debug mode with enhanced logging."""
    global range_image_logger

    if enabled:
        range_image_logger = RangeImageLogger(level=level)
        range_image_logger.info("🚀 RANGE IMAGE DEBUG MODE ACTIVATED")
        range_image_logger.info(
            "🔍 Enhanced logging enabled for range image operations"
        )
    else:
        range_image_logger.info("🛑 RANGE IMAGE DEBUG MODE DEACTIVATED")


def log_range_config(config_data: dict, index: int = 0) -> None:
    """Log range configuration with detailed formatting."""
    range_image_logger.info(
        f"📋 RANGE CONFIG [{index}]:\n"
        f"   🏷️  Field Name: {config_data.get('field_name', 'N/A')}\n"
        f"   📊 Sheet Name: {config_data.get('sheet_name', 'N/A')}\n"
        f"   📍 Range: {config_data.get('range', 'N/A')}\n"
        f"   🎨 Format: {config_data.get('output_format', 'png')}\n"
        f"   📐 DPI: {config_data.get('dpi', 150)}\n"
        f"   📏 Dimensions: {config_data.get('width', 'auto')} x {config_data.get('height', 'auto')}"
    )


def log_graph_api_status(
    client_id: str, status: str, details: Optional[str] = None
) -> None:
    """Log Graph API connection status."""
    masked_client_id = f"{client_id[:8]}..." if len(client_id) > 8 else client_id

    status_emoji = {
        "connecting": "🔄",
        "connected": "✅",
        "failed": "❌",
        "authenticating": "🔐",
        "authenticated": "🔓",
    }.get(status, "❓")

    message = f"{status_emoji} GRAPH API {status.upper()}: {masked_client_id}"
    if details:
        message += f"\n   📝 Details: {details}"

    if status in ["failed", "error"]:
        range_image_logger.error(message)
    elif status in ["connecting", "authenticating"]:
        range_image_logger.info(message)
    else:
        range_image_logger.info(message)


def log_range_export_progress(
    current: int, total: int, field_name: str, status: str
) -> None:
    """Log range export progress."""
    progress_bar = "█" * int((current / total) * 20) + "░" * (
        20 - int((current / total) * 20)
    )
    percentage = int((current / total) * 100)

    status_emoji = {
        "processing": "⚙️",
        "success": "✅",
        "failed": "❌",
        "uploading": "⬆️",
        "rendering": "🎨",
    }.get(status, "❓")

    range_image_logger.info(
        f"{status_emoji} EXPORT PROGRESS [{current}/{total}] {percentage}%\n"
        f"   {progress_bar}\n"
        f"   🏷️  Current: {field_name}\n"
        f"   🎯 Status: {status.upper()}"
    )


def log_range_validation_result(
    sheet_name: str,
    range_str: str,
    is_valid: bool,
    has_data: bool = None,
    error: str = None,
) -> None:
    """Log range validation results."""
    if is_valid:
        data_status = (
            "✅ HAS DATA"
            if has_data
            else "⚠️ EMPTY" if has_data is False else "❓ UNKNOWN"
        )
        range_image_logger.info(
            f"✅ RANGE VALIDATION PASSED\n"
            f"   📊 Sheet: {sheet_name}\n"
            f"   📍 Range: {range_str}\n"
            f"   📦 Data Status: {data_status}"
        )
    else:
        range_image_logger.error(
            f"❌ RANGE VALIDATION FAILED\n"
            f"   📊 Sheet: {sheet_name}\n"
            f"   📍 Range: {range_str}\n"
            f"   💥 Error: {error or 'Unknown error'}"
        )
