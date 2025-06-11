"""Configuration management system with JSON schema validation and enhanced image extraction settings."""

import json
import os
import datetime
from typing import Any, Dict, List, Optional
import logging
from dotenv import load_dotenv

from .utils.exceptions import ConfigurationError
from .utils.validation import validate_config_structure

logger = logging.getLogger(__name__)


class ConfigManager:
    """Manages application configuration with validation and environment support."""

    def __init__(self, config_dir: str = "config") -> None:
        """Initialize configuration manager."""
        self.config_dir = config_dir
        self._config_cache: Dict[str, Dict[str, Any]] = {}
        self._load_environment()

    def _load_environment(self) -> None:
        """Load environment variables from .env files."""
        try:
            # Load from .env file if it exists
            env_file = os.path.join(os.getcwd(), '.env')
            if os.path.exists(env_file):
                load_dotenv(env_file)
                logger.debug("Loaded environment from .env file")

            # Load environment-specific file
            env = os.getenv('ENVIRONMENT', 'development')
            env_specific_file = os.path.join(self.config_dir, f"{env}.env")

            if os.path.exists(env_specific_file):
                load_dotenv(env_specific_file)
                logger.debug(f"Loaded environment from {env_specific_file}")

        except Exception as e:
            logger.warning(f"Failed to load environment configuration: {e}")

    def get_default_config(self) -> Dict[str, Any]:
        """Get default configuration for Excel data extraction with enhanced image settings."""
        return {
            "version": "1.0",
            "sheet_configs": {
                "Order Form": {
                    "subtables": [
                        {
                            "name": "client_info",
                            "type": "key_value_pairs",
                            "header_search": {
                                "method": "contains_text",
                                "text": "Client",
                                "column": "A",
                                "search_range": "A1:A10"
                            },
                            "data_extraction": {
                                "orientation": "horizontal",
                                "headers_row_offset": 0,
                                "data_row_offset": 1,
                                "max_columns": 6,
                                "column_mappings": {
                                    "Client": "client_name",
                                    "Word Or Image": "search_type",
                                    "G&S Classes": "gs_classes",
                                    "SIC": "sic_code",
                                    "Nature of business": "business_nature",
                                    "Designated Countries": "countries"
                                }
                            }
                        },
                        {
                            "name": "word_search",
                            "type": "table",
                            "header_search": {
                                "method": "contains_text",
                                "text": "Word",
                                "column": "A",
                                "search_range": "A1:A20"
                            },
                            "data_extraction": {
                                "orientation": "vertical",
                                "headers_row_offset": 0,
                                "data_row_offset": 1,
                                "max_columns": 3,
                                "max_rows": 10,
                                "column_mappings": {
                                    "Word": "word",
                                    "Search Criteria": "search_criteria",
                                    "Remarks": "remarks"
                                }
                            }
                        },
                        {
                            "name": "image_search",
                            "type": "table",
                            "header_search": {
                                "method": "contains_text",
                                "text": "Image",
                                "column": "A",
                                "search_range": "A10:A30"
                            },
                            "data_extraction": {
                                "orientation": "vertical",
                                "headers_row_offset": 0,
                                "data_row_offset": 1,
                                "max_columns": 3,
                                "max_rows": 10,
                                "column_mappings": {
                                    "Image": "image",
                                    "Search Criteria": "search_criteria",
                                    "Image Class.Division.Subdivision": "image_classification"
                                }
                            }
                        }
                    ]
                }
            },
            "global_settings": {
                "normalize_keys": True,
                "temp_file_cleanup": {
                    "enabled": True,
                    "delay_seconds": 300,
                    "keep_on_error": True,
                    "development_mode": False
                },
                "image_extraction": {
                    "enabled": True,
                    "preserve_format": True,
                    "extract_position": True,
                    "fallback_format": "png",
                    "max_size_mb": 10,
                    "supported_formats": ["png", "jpg", "jpeg", "gif", "webp"],
                    "position_matching": {
                        "enabled": True,
                        "confidence_threshold": 0.3,
                        "prefer_position_over_index": True
                    },
                    "debug_mode": {
                        "save_position_info": True,
                        "log_anchor_details": True,
                        "create_summary": True
                    }
                },
                "validation": {
                    "strict_mode": False,
                    "allow_empty_values": True,
                    "required_fields": [],
                    "validate_image_placeholders": True
                },
                "powerpoint_processing": {
                    "image_placeholder_patterns": [
                        r'\{\{.*image.*\}\}',
                        r'\{\{.*img.*\}\}',
                        r'\{\{.*photo.*\}\}',
                        r'\{\{.*picture.*\}\}'
                    ],
                    "auto_resize_images": True,
                    "maintain_aspect_ratio": True,
                    "default_image_size": {
                        "width_inches": 2.0,
                        "height_inches": 2.0
                    }
                }
            }
        }

    def load_config(self, config_name: str = "default_config") -> Dict[str, Any]:
        """Load configuration from file with caching."""
        if config_name in self._config_cache:
            return self._config_cache[config_name]

        try:
            config_file = os.path.join(self.config_dir, f"{config_name}.json")

            if not os.path.exists(config_file):
                if config_name == "default_config":
                    # Return default configuration if file doesn't exist
                    config = self.get_default_config()
                    self._config_cache[config_name] = config
                    return config
                else:
                    raise ConfigurationError(f"Configuration file not found: {config_file}")

            with open(config_file, 'r', encoding='utf-8') as file:
                config = json.load(file)

            # Validate configuration structure
            validate_config_structure(config)

            # Apply environment overrides
            config = self._apply_environment_overrides(config)

            # Cache the configuration
            self._config_cache[config_name] = config

            logger.info(f"Loaded configuration: {config_name}")
            return config

        except (json.JSONDecodeError, FileNotFoundError) as e:
            raise ConfigurationError(f"Failed to load configuration '{config_name}': {e}")
        except Exception as e:
            raise ConfigurationError(f"Configuration error for '{config_name}': {e}")

    def save_config(self, config: Dict[str, Any], config_name: str) -> None:
        """Save configuration to file."""
        try:
            # Validate configuration before saving
            validate_config_structure(config)

            # Ensure config directory exists
            os.makedirs(self.config_dir, exist_ok=True)

            config_file = os.path.join(self.config_dir, f"{config_name}.json")

            with open(config_file, 'w', encoding='utf-8') as file:
                json.dump(config, file, indent=2, ensure_ascii=False)

            # Update cache
            self._config_cache[config_name] = config

            logger.info(f"Saved configuration: {config_name}")

        except Exception as e:
            raise ConfigurationError(f"Failed to save configuration '{config_name}': {e}")

    def _apply_environment_overrides(self, config: Dict[str, Any]) -> Dict[str, Any]:
        """Apply environment variable overrides to configuration."""
        try:
            # Override global settings from environment
            if 'global_settings' in config:
                global_settings = config['global_settings']

                # Override temp file cleanup settings
                if 'temp_file_cleanup' in global_settings:
                    cleanup_config = global_settings['temp_file_cleanup']

                    cleanup_config['enabled'] = self._get_env_bool(
                        'CLEANUP_TEMP_FILES',
                        cleanup_config.get('enabled', True)
                    )

                    cleanup_config['delay_seconds'] = self._get_env_int(
                        'TEMP_FILE_RETENTION_SECONDS',
                        cleanup_config.get('delay_seconds', 300)
                    )

                    cleanup_config['development_mode'] = self._get_env_bool(
                        'DEVELOPMENT_MODE',
                        cleanup_config.get('development_mode', False)
                    )

                # Override image extraction settings
                if 'image_extraction' in global_settings:
                    image_config = global_settings['image_extraction']

                    image_config['enabled'] = self._get_env_bool(
                        'IMAGE_EXTRACTION_ENABLED',
                        image_config.get('enabled', True)
                    )

                    image_config['extract_position'] = self._get_env_bool(
                        'EXTRACT_IMAGE_POSITION',
                        image_config.get('extract_position', True)
                    )

                    image_config['max_size_mb'] = self._get_env_int(
                        'MAX_IMAGE_SIZE_MB',
                        image_config.get('max_size_mb', 10)
                    )

                    # Override position matching settings
                    if 'position_matching' in image_config:
                        position_config = image_config['position_matching']

                        position_config['enabled'] = self._get_env_bool(
                            'POSITION_MATCHING_ENABLED',
                            position_config.get('enabled', True)
                        )

                        position_config['confidence_threshold'] = self._get_env_float(
                            'POSITION_MATCHING_THRESHOLD',
                            position_config.get('confidence_threshold', 0.3)
                        )

                    # Override debug mode settings
                    if 'debug_mode' in image_config:
                        debug_config = image_config['debug_mode']

                        debug_config['save_position_info'] = self._get_env_bool(
                            'SAVE_IMAGE_POSITION_INFO',
                            debug_config.get('save_position_info', True)
                        )

                        debug_config['log_anchor_details'] = self._get_env_bool(
                            'LOG_IMAGE_ANCHOR_DETAILS',
                            debug_config.get('log_anchor_details', True)
                        )

                # Override normalization setting
                global_settings['normalize_keys'] = self._get_env_bool(
                    'NORMALIZE_COLUMN_KEYS',
                    global_settings.get('normalize_keys', True)
                )

                # Override PowerPoint processing settings
                if 'powerpoint_processing' in global_settings:
                    pptx_config = global_settings['powerpoint_processing']

                    pptx_config['auto_resize_images'] = self._get_env_bool(
                        'AUTO_RESIZE_IMAGES',
                        pptx_config.get('auto_resize_images', True)
                    )

                    pptx_config['maintain_aspect_ratio'] = self._get_env_bool(
                        'MAINTAIN_ASPECT_RATIO',
                        pptx_config.get('maintain_aspect_ratio', True)
                    )

            return config

        except Exception as e:
            logger.warning(f"Failed to apply environment overrides: {e}")
            return config

    def _get_env_bool(self, key: str, default: bool) -> bool:
        """Get boolean value from environment variable."""
        value = os.getenv(key, str(default)).lower()
        return value in ('true', '1', 'yes', 'on')

    def _get_env_int(self, key: str, default: int) -> int:
        """Get integer value from environment variable."""
        try:
            return int(os.getenv(key, str(default)))
        except (ValueError, TypeError):
            return default

    def _get_env_float(self, key: str, default: float) -> float:
        """Get float value from environment variable."""
        try:
            return float(os.getenv(key, str(default)))
        except (ValueError, TypeError):
            return default

    def _get_env_str(self, key: str, default: str) -> str:
        """Get string value from environment variable."""
        return os.getenv(key, default)

    def _get_env_list(self, key: str, default: List[str], separator: str = ',') -> List[str]:
        """Get list value from environment variable."""
        value = os.getenv(key)
        if value:
            return [item.strip() for item in value.split(separator) if item.strip()]
        return default

    def get_app_config(self) -> Dict[str, Any]:
        """Get application-wide configuration settings."""
        return {
            'development_mode': self._get_env_bool('DEVELOPMENT_MODE', False),
            'log_level': self._get_env_str('LOG_LEVEL', 'INFO'),
            'api_key': self._get_env_str('API_KEY', ''),
            'max_file_size_mb': self._get_env_int('MAX_FILE_SIZE_MB', 50),
            'allowed_extensions': self._get_env_list('ALLOWED_EXTENSIONS', ['xlsx', 'pptx']),
            'temp_directory': self._get_env_str('TEMP_DIRECTORY', '/tmp'),
            'flask_config': {
                'host': self._get_env_str('FLASK_HOST', '0.0.0.0'),
                'port': self._get_env_int('FLASK_PORT', 5000),
                'debug': self._get_env_bool('FLASK_DEBUG', False)
            },
            'google_cloud': {
                'project_id': self._get_env_str('GOOGLE_CLOUD_PROJECT', ''),
                'bucket_name': self._get_env_str('GOOGLE_CLOUD_BUCKET', ''),
                'function_name': self._get_env_str('GOOGLE_CLOUD_FUNCTION', 'excel-pptx-merger')
            },
            'zoho_workdrive': {
                'client_id': self._get_env_str('ZOHO_CLIENT_ID', ''),
                'client_secret': self._get_env_str('ZOHO_CLIENT_SECRET', ''),
                'refresh_token': self._get_env_str('ZOHO_REFRESH_TOKEN', '')
            },
            'image_processing': {
                'max_concurrent_extractions': self._get_env_int('MAX_CONCURRENT_IMAGE_EXTRACTIONS', 5),
                'timeout_seconds': self._get_env_int('IMAGE_PROCESSING_TIMEOUT', 30),
                'quality_optimization': self._get_env_bool('OPTIMIZE_IMAGE_QUALITY', True)
            }
        }

    def get_image_extraction_config(self) -> Dict[str, Any]:
        """Get specific configuration for image extraction."""
        default_config = self.get_default_config()
        return default_config.get('global_settings', {}).get('image_extraction', {})

    def get_powerpoint_config(self) -> Dict[str, Any]:
        """Get specific configuration for PowerPoint processing."""
        default_config = self.get_default_config()
        return default_config.get('global_settings', {}).get('powerpoint_processing', {})

    def merge_configs(self, base_config: Dict[str, Any], override_config: Dict[str, Any]) -> Dict[str, Any]:
        """Merge two configurations with override taking precedence."""
        try:
            merged = base_config.copy()

            for key, value in override_config.items():
                if key in merged and isinstance(merged[key], dict) and isinstance(value, dict):
                    # Recursively merge dictionaries
                    merged[key] = self.merge_configs(merged[key], value)
                else:
                    # Override value
                    merged[key] = value

            return merged

        except Exception as e:
            raise ConfigurationError(f"Failed to merge configurations: {e}")

    def validate_runtime_config(self, config: Dict[str, Any]) -> None:
        """Validate configuration at runtime with additional checks."""
        try:
            # Basic structure validation
            validate_config_structure(config)

            # Additional runtime validations
            if 'sheet_configs' in config:
                for sheet_name, sheet_config in config['sheet_configs'].items():
                    if not sheet_name.strip():
                        raise ConfigurationError("Sheet name cannot be empty")

                    if 'subtables' not in sheet_config:
                        raise ConfigurationError(f"Sheet '{sheet_name}' missing subtables configuration")

                    for subtable in sheet_config['subtables']:
                        self._validate_subtable_config(subtable)

            # Validate global settings
            if 'global_settings' in config:
                self._validate_global_settings(config['global_settings'])

        except Exception as e:
            raise ConfigurationError(f"Runtime configuration validation failed: {e}")

    def _validate_subtable_config(self, subtable: Dict[str, Any]) -> None:
        """Validate individual subtable configuration."""
        required_fields = ['name', 'type', 'header_search', 'data_extraction']

        for field in required_fields:
            if field not in subtable:
                raise ConfigurationError(f"Subtable missing required field: {field}")

        # Validate subtable type
        valid_types = ['key_value_pairs', 'table']
        if subtable['type'] not in valid_types:
            raise ConfigurationError(f"Invalid subtable type: {subtable['type']}")

        # Validate header search configuration
        header_search = subtable['header_search']
        if 'method' not in header_search:
            raise ConfigurationError("Header search missing method")

        valid_methods = ['contains_text', 'exact_match', 'regex']
        if header_search['method'] not in valid_methods:
            raise ConfigurationError(f"Invalid header search method: {header_search['method']}")

    def _validate_global_settings(self, settings: Dict[str, Any]) -> None:
        """Validate global settings configuration."""
        if 'temp_file_cleanup' in settings:
            cleanup = settings['temp_file_cleanup']
            if 'delay_seconds' in cleanup and cleanup['delay_seconds'] < 0:
                raise ConfigurationError("Temp file cleanup delay cannot be negative")

        if 'image_extraction' in settings:
            img_config = settings['image_extraction']
            if 'max_size_mb' in img_config and img_config['max_size_mb'] <= 0:
                raise ConfigurationError("Image max size must be positive")

            # Validate supported formats
            if 'supported_formats' in img_config:
                valid_formats = ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'tiff']
                for fmt in img_config['supported_formats']:
                    if fmt.lower() not in valid_formats:
                        raise ConfigurationError(f"Unsupported image format: {fmt}")

            # Validate position matching configuration
            if 'position_matching' in img_config:
                pos_config = img_config['position_matching']
                if 'confidence_threshold' in pos_config:
                    threshold = pos_config['confidence_threshold']
                    if not (0.0 <= threshold <= 1.0):
                        raise ConfigurationError("Position matching confidence threshold must be between 0.0 and 1.0")

        if 'powerpoint_processing' in settings:
            pptx_config = settings['powerpoint_processing']

            # Validate default image size
            if 'default_image_size' in pptx_config:
                size_config = pptx_config['default_image_size']
                if 'width_inches' in size_config and size_config['width_inches'] <= 0:
                    raise ConfigurationError("Default image width must be positive")
                if 'height_inches' in size_config and size_config['height_inches'] <= 0:
                    raise ConfigurationError("Default image height must be positive")

            # Validate placeholder patterns
            if 'image_placeholder_patterns' in pptx_config:
                import re
                for pattern in pptx_config['image_placeholder_patterns']:
                    try:
                        re.compile(pattern)
                    except re.error as e:
                        raise ConfigurationError(f"Invalid regex pattern for image placeholder: {pattern} - {e}")
    
    def clear_cache(self) -> None:
        """Clear configuration cache."""
        self._config_cache.clear()
        logger.debug("Configuration cache cleared")
    
    def get_cached_configs(self) -> List[str]:
        """Get list of cached configuration names."""
        return list(self._config_cache.keys())