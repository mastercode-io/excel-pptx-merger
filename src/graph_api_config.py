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
        """Load configuration from environment file or environment variables with detailed logging."""
        logger.info("🔧 Starting GraphAPIConfig._load_config()...")
        logger.info(f"🔧 Initial config_file parameter: {self.config_file}")
        
        # Try to load from file first
        if self.config_file and os.path.exists(self.config_file):
            logger.info(f"📁 Loading from specified config file: {self.config_file}")
            self._load_from_file()
        else:
            logger.info("📁 No specific config file provided or file doesn't exist")
            # Try default locations
            default_paths = ["config/graph_api.env", "graph_api.env", ".env"]
            logger.info(f"📁 Trying default config file paths: {default_paths}")

            for path in default_paths:
                logger.info(f"📁 Checking path: {path}")
                if os.path.exists(path):
                    logger.info(f"📁 Found config file at: {path}")
                    self.config_file = path
                    self._load_from_file()
                    break
                else:
                    logger.info(f"📁 Path does not exist: {path}")
            else:
                logger.warning("📁 No config file found in any default location")

        # Load from environment variables (takes precedence)
        logger.info("🔧 Loading from environment variables (takes precedence over file)...")
        self._load_from_env()

        # Set default settings
        logger.info("🔧 Setting default settings...")
        self._set_defaults()
        
        logger.info("🔧 GraphAPIConfig._load_config() completed")

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
        """Load configuration from environment variables with detailed logging."""
        logger.info("🔧 Loading credentials from environment variables...")
        
        client_id = os.getenv("GRAPH_CLIENT_ID", "")
        client_secret = os.getenv("GRAPH_CLIENT_SECRET", "")
        tenant_id = os.getenv("GRAPH_TENANT_ID", "")
        
        # Log environment variable status (masked for security)
        logger.info(f"🔑 GRAPH_CLIENT_ID: {'SET (' + client_id[:8] + '...)' if client_id else 'NOT SET'}")
        logger.info(f"🔑 GRAPH_CLIENT_SECRET: {'SET (' + client_secret[:4] + '...)' if client_secret else 'NOT SET'}")
        logger.info(f"🔑 GRAPH_TENANT_ID: {'SET (' + tenant_id[:8] + '...)' if tenant_id else 'NOT SET'}")
        
        self._credentials = {
            "client_id": client_id,
            "client_secret": client_secret,
            "tenant_id": tenant_id,
        }
        
        logger.info(f"🔧 Credentials loaded from environment: {len([k for k, v in self._credentials.items() if v])} of 3 set")

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
        """Check if Graph API is properly configured with detailed logging."""
        logger.info("🔧 Checking GraphAPIConfig.is_configured()...")
        
        client_id = self._credentials.get("client_id")
        client_secret = self._credentials.get("client_secret")
        tenant_id = self._credentials.get("tenant_id")
        
        logger.info(f"🔧 client_id present: {bool(client_id)}")
        logger.info(f"🔧 client_secret present: {bool(client_secret)}")
        logger.info(f"🔧 tenant_id present: {bool(tenant_id)}")
        
        is_configured = all([client_id, client_secret, tenant_id])
        logger.info(f"🔧 Overall is_configured result: {is_configured}")
        
        if not is_configured:
            logger.error("❌ GraphAPIConfig is not configured - missing required credentials")
            if not client_id:
                logger.error("❌ Missing client_id")
            if not client_secret:
                logger.error("❌ Missing client_secret")  
            if not tenant_id:
                logger.error("❌ Missing tenant_id")
        
        return is_configured

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
    """Get Graph API credentials if available with detailed diagnostics.

    Args:
        tenant_id: Optional tenant ID from configuration to override environment variable
    """
    logger.info("🔧 get_graph_api_credentials() called")
    logger.info(f"🔧 Input tenant_id parameter: {tenant_id}")
    
    try:
        logger.info("🔧 Loading Graph API config...")
        config = load_graph_api_config()
        logger.info(f"🔧 Config loaded successfully: {type(config)}")
        
        logger.info("🔧 Checking if config is configured...")
        is_configured = config.is_configured()
        logger.info(f"🔧 Config is_configured(): {is_configured}")
        
        if is_configured:
            logger.info("✅ Config is configured - getting credentials...")
            credentials = config.get_credentials()
            logger.info(f"✅ Raw credentials obtained: {type(credentials)}")
            logger.info(f"✅ Credential keys: {list(credentials.keys()) if credentials else 'None'}")
            
            # Log individual credential availability (masked)
            if credentials:
                for key in ['client_id', 'client_secret', 'tenant_id']:
                    value = credentials.get(key, '')
                    if value:
                        masked_value = value[:4] + '...' if len(value) > 4 else value
                        logger.info(f"✅ {key}: SET ({masked_value})")
                    else:
                        logger.error(f"❌ {key}: NOT SET")

            # Override tenant_id if provided from config
            if tenant_id:
                original_tenant = credentials.get("tenant_id", "")
                credentials["tenant_id"] = tenant_id
                logger.info(f"🔄 Overrode tenant_id: '{original_tenant}' -> '{tenant_id}'")
            
            logger.info("✅ Returning credentials")
            return credentials
        else:
            logger.error("❌ Graph API not configured - credentials unavailable")
            logger.error("❌ This means environment variables or config file are missing")
            
            # Get detailed configuration status
            summary = config.get_config_summary()
            logger.error(f"❌ Configuration summary:\n{summary}")
            
            return None
            
    except Exception as e:
        logger.error(f"❌ Exception in get_graph_api_credentials(): {e}")
        logger.error(f"❌ Exception type: {type(e)}")
        import traceback
        logger.error(f"❌ Traceback: {traceback.format_exc()}")
        return None
