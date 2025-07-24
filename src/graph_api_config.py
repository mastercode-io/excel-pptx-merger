"""Configuration management for Microsoft Graph API."""

import os
import logging
from typing import Dict, Optional
from pathlib import Path

logger = logging.getLogger(__name__)


class GraphAPIConfig:
    """Configuration manager for Graph API credentials and settings."""

    def __init__(self, config_file: Optional[str] = None):
        """Initialize Graph API configuration.

        Args:
            config_file: Path to environment file with Graph API credentials
        """
        self.config_file = config_file
        self._credentials = None
        self._settings = None
        self._load_config()

    def _load_config(self) -> None:
        """Load configuration from environment file or environment variables."""
        # Try to load from file first
        if self.config_file and os.path.exists(self.config_file):
            self._load_from_file()
        else:
            # Try default locations
            default_paths = ["config/graph_api.env", "graph_api.env", ".env"]

            for path in default_paths:
                if os.path.exists(path):
                    self.config_file = path
                    self._load_from_file()
                    break

        # Load from environment variables (takes precedence)
        self._load_from_env()

        # Set default settings
        self._set_defaults()

    def _load_from_file(self) -> None:
        """Load configuration from environment file."""
        try:
            with open(self.config_file, "r") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        key, value = line.split("=", 1)
                        os.environ[key.strip()] = value.strip()

            logger.info(f"Loaded Graph API configuration from: {self.config_file}")
        except Exception as e:
            logger.warning(
                f"Failed to load Graph API config file {self.config_file}: {e}"
            )

    def _load_from_env(self) -> None:
        """Load configuration from environment variables."""
        self._credentials = {
            "client_id": os.getenv("GRAPH_CLIENT_ID", ""),
            "client_secret": os.getenv("GRAPH_CLIENT_SECRET", ""),
            "tenant_id": os.getenv("GRAPH_TENANT_ID", ""),
        }

        self._settings = {
            "timeout": int(os.getenv("GRAPH_API_TIMEOUT", "60")),
            "retry_attempts": int(os.getenv("GRAPH_API_RETRY_ATTEMPTS", "3")),
            "retry_delay": int(os.getenv("GRAPH_API_RETRY_DELAY", "2")),
            "max_range_cells": int(os.getenv("RANGE_EXPORT_MAX_CELLS", "10000")),
            "default_dpi": int(os.getenv("RANGE_EXPORT_DEFAULT_DPI", "150")),
            "default_format": os.getenv("RANGE_EXPORT_DEFAULT_FORMAT", "png"),
            "temp_cleanup": os.getenv("RANGE_EXPORT_TEMP_CLEANUP", "true").lower()
            == "true",
        }

    def _set_defaults(self) -> None:
        """Set default values for missing configuration."""
        if not self._credentials:
            self._credentials = {"client_id": "", "client_secret": "", "tenant_id": ""}

        if not self._settings:
            self._settings = {
                "timeout": 60,
                "retry_attempts": 3,
                "retry_delay": 2,
                "max_range_cells": 10000,
                "default_dpi": 150,
                "default_format": "png",
                "temp_cleanup": True,
            }

    def get_credentials(self) -> Dict[str, str]:
        """Get Graph API credentials."""
        return self._credentials.copy()

    def get_settings(self) -> Dict[str, any]:
        """Get Graph API settings."""
        return self._settings.copy()

    def is_configured(self) -> bool:
        """Check if Graph API is properly configured."""
        return all(
            [
                self._credentials.get("client_id"),
                self._credentials.get("client_secret"),
                self._credentials.get("tenant_id"),
            ]
        )

    def validate_config(self) -> tuple[bool, list[str]]:
        """Validate configuration and return any errors."""
        errors = []

        # Check required credentials
        required_creds = ["client_id", "client_secret", "tenant_id"]
        for cred in required_creds:
            if not self._credentials.get(cred):
                errors.append(f"Missing required credential: {cred}")

        # Validate settings
        if self._settings.get("timeout", 0) <= 0:
            errors.append("Timeout must be positive")

        if self._settings.get("retry_attempts", 0) < 0:
            errors.append("Retry attempts must be non-negative")

        if (
            self._settings.get("default_dpi", 0) < 72
            or self._settings.get("default_dpi", 0) > 600
        ):
            errors.append("Default DPI must be between 72 and 600")

        valid_formats = ["png", "jpg", "jpeg"]
        if self._settings.get("default_format", "").lower() not in valid_formats:
            errors.append(f"Default format must be one of: {valid_formats}")

        return len(errors) == 0, errors

    def get_config_summary(self) -> str:
        """Get a summary of the current configuration."""
        is_valid, errors = self.validate_config()
        status = "✓ Valid" if is_valid else "✗ Invalid"

        summary = f"""Graph API Configuration Summary:
Status: {status}
Config File: {self.config_file or 'None (using environment variables)'}
Client ID: {'Set' if self._credentials.get('client_id') else 'Missing'}
Client Secret: {'Set' if self._credentials.get('client_secret') else 'Missing'}
Tenant ID: {'Set' if self._credentials.get('tenant_id') else 'Missing'}
Settings: {self._settings}"""

        if errors:
            summary += f"\nErrors: {errors}"

        return summary


def load_graph_api_config(config_file: Optional[str] = None) -> GraphAPIConfig:
    """Load Graph API configuration from file or environment."""
    return GraphAPIConfig(config_file)


def get_graph_api_credentials(
    tenant_id: Optional[str] = None,
) -> Optional[Dict[str, str]]:
    """Get Graph API credentials if available.

    Args:
        tenant_id: Optional tenant ID from configuration to override environment variable
    """
    config = load_graph_api_config()

    if config.is_configured():
        credentials = config.get_credentials()

        # Override tenant_id if provided from config
        if tenant_id:
            credentials["tenant_id"] = tenant_id
            logger.info("Using tenant_id from configuration")

        return credentials
    else:
        logger.warning("Graph API not configured - range image export will be disabled")
        return None
