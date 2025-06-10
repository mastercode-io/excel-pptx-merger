"""Configuration management system with JSON schema validation."""

import json
import os
from typing import Any, Dict, Optional
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
        """Get default configuration for Excel data extraction."""
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
                                    "Word Or Image": "search_type"
                                }
                            }
                        },
                        {
                            "name": "trademark_classes",
                            "type": "table",
                            "header_search": {
                                "method": "contains_text",
                                "text": "Class",
                                "column": "A",
                                "search_range": "A10:A30"
                            },
                            "data_extraction": {
                                "orientation": "vertical",
                                "headers_row_offset": 0,
                                "data_row_offset": 1,
                                "max_columns": 5,
                                "max_rows": 20,
                                "column_mappings": {
                                    "Class": "class_number",
                                    "Description": "class_description",
                                    "Goods/Services": "goods_services"
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
                    "formats": ["png", "jpg", "jpeg"],
                    "max_size_mb": 10
                },
                "validation": {
                    "strict_mode": False,
                    "allow_empty_values": True,
                    "required_fields": []
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
                
                # Override normalization setting
                global_settings['normalize_keys'] = self._get_env_bool(
                    'NORMALIZE_COLUMN_KEYS',
                    global_settings.get('normalize_keys', True)
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
    
    def _get_env_str(self, key: str, default: str) -> str:
        """Get string value from environment variable."""
        return os.getenv(key, default)
    
    def get_app_config(self) -> Dict[str, Any]:
        """Get application-wide configuration settings."""
        return {
            'development_mode': self._get_env_bool('DEVELOPMENT_MODE', False),
            'log_level': self._get_env_str('LOG_LEVEL', 'INFO'),
            'api_key': self._get_env_str('API_KEY', ''),
            'max_file_size_mb': self._get_env_int('MAX_FILE_SIZE_MB', 50),
            'allowed_extensions': self._get_env_str('ALLOWED_EXTENSIONS', 'xlsx,pptx').split(','),
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
            }
        }
    
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
    
    def clear_cache(self) -> None:
        """Clear configuration cache."""
        self._config_cache.clear()
        logger.debug("Configuration cache cleared")
    
    def get_cached_configs(self) -> List[str]:
        """Get list of cached configuration names."""
        return list(self._config_cache.keys())