"""Tests for configuration manager module."""

import json
import os
import pytest
import tempfile
from unittest.mock import patch, mock_open

from src.config_manager import ConfigManager
from src.utils.exceptions import ConfigurationError, ValidationError


class TestConfigManager:
    """Test cases for ConfigManager class."""
    
    @pytest.fixture
    def temp_config_dir(self):
        """Create temporary config directory for testing."""
        with tempfile.TemporaryDirectory() as temp_dir:
            yield temp_dir
    
    @pytest.fixture
    def sample_config(self):
        """Sample configuration for testing."""
        return {
            "version": "1.0",
            "sheet_configs": {
                "TestSheet": {
                    "subtables": [
                        {
                            "name": "test_data",
                            "type": "key_value_pairs",
                            "header_search": {
                                "method": "contains_text",
                                "text": "Test",
                                "column": "A",
                                "search_range": "A1:A5"
                            },
                            "data_extraction": {
                                "orientation": "horizontal",
                                "headers_row_offset": 0,
                                "data_row_offset": 1,
                                "max_columns": 3
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
                }
            }
        }
    
    def test_init(self, temp_config_dir):
        """Test ConfigManager initialization."""
        manager = ConfigManager(temp_config_dir)
        assert manager.config_dir == temp_config_dir
        assert isinstance(manager._config_cache, dict)
    
    def test_get_default_config(self, temp_config_dir):
        """Test getting default configuration."""
        manager = ConfigManager(temp_config_dir)
        config = manager.get_default_config()
        
        assert "version" in config
        assert "sheet_configs" in config
        assert "global_settings" in config
        assert config["version"] == "1.0"
        assert isinstance(config["sheet_configs"], dict)
        assert isinstance(config["global_settings"], dict)
    
    def test_load_config_default(self, temp_config_dir):
        """Test loading default configuration when file doesn't exist."""
        manager = ConfigManager(temp_config_dir)
        config = manager.load_config("default_config")
        
        # Should return default config since file doesn't exist
        assert "version" in config
        assert "sheet_configs" in config
        assert "global_settings" in config
    
    def test_load_config_from_file(self, temp_config_dir, sample_config):
        """Test loading configuration from file."""
        # Create config file
        config_file = os.path.join(temp_config_dir, "test_config.json")
        with open(config_file, 'w') as f:
            json.dump(sample_config, f)
        
        manager = ConfigManager(temp_config_dir)
        config = manager.load_config("test_config")
        
        assert config["version"] == "1.0"
        assert "TestSheet" in config["sheet_configs"]
        assert config["global_settings"]["normalize_keys"] is True
    
    def test_load_config_invalid_json(self, temp_config_dir):
        """Test loading configuration with invalid JSON."""
        # Create invalid JSON file
        config_file = os.path.join(temp_config_dir, "invalid_config.json")
        with open(config_file, 'w') as f:
            f.write("{ invalid json }")
        
        manager = ConfigManager(temp_config_dir)
        with pytest.raises(ConfigurationError):
            manager.load_config("invalid_config")
    
    def test_save_config(self, temp_config_dir, sample_config):
        """Test saving configuration to file."""
        manager = ConfigManager(temp_config_dir)
        manager.save_config(sample_config, "saved_config")
        
        # Verify file was created
        config_file = os.path.join(temp_config_dir, "saved_config.json")
        assert os.path.exists(config_file)
        
        # Verify content
        with open(config_file, 'r') as f:
            loaded_config = json.load(f)
        
        assert loaded_config == sample_config
    
    def test_save_config_invalid(self, temp_config_dir):
        """Test saving invalid configuration."""
        invalid_config = {
            "version": "1.0",
            # Missing required fields
        }
        
        manager = ConfigManager(temp_config_dir)
        with pytest.raises(ConfigurationError):
            manager.save_config(invalid_config, "invalid_config")
    
    @patch.dict(os.environ, {
        'CLEANUP_TEMP_FILES': 'false',
        'TEMP_FILE_RETENTION_SECONDS': '600',
        'DEVELOPMENT_MODE': 'true'
    })
    def test_apply_environment_overrides(self, temp_config_dir, sample_config):
        """Test applying environment variable overrides."""
        manager = ConfigManager(temp_config_dir)
        config = manager._apply_environment_overrides(sample_config)
        
        cleanup_config = config["global_settings"]["temp_file_cleanup"]
        assert cleanup_config["enabled"] is False  # From CLEANUP_TEMP_FILES
        assert cleanup_config["delay_seconds"] == 600  # From TEMP_FILE_RETENTION_SECONDS
        assert cleanup_config["development_mode"] is True  # From DEVELOPMENT_MODE
    
    def test_get_env_bool(self, temp_config_dir):
        """Test getting boolean values from environment."""
        manager = ConfigManager(temp_config_dir)
        
        # Test various true values
        with patch.dict(os.environ, {'TEST_BOOL': 'true'}):
            assert manager._get_env_bool('TEST_BOOL', False) is True
        
        with patch.dict(os.environ, {'TEST_BOOL': '1'}):
            assert manager._get_env_bool('TEST_BOOL', False) is True
        
        with patch.dict(os.environ, {'TEST_BOOL': 'yes'}):
            assert manager._get_env_bool('TEST_BOOL', False) is True
        
        # Test false values
        with patch.dict(os.environ, {'TEST_BOOL': 'false'}):
            assert manager._get_env_bool('TEST_BOOL', True) is False
        
        with patch.dict(os.environ, {'TEST_BOOL': '0'}):
            assert manager._get_env_bool('TEST_BOOL', True) is False
        
        # Test default when not set
        with patch.dict(os.environ, {}, clear=True):
            assert manager._get_env_bool('TEST_BOOL', True) is True
    
    def test_get_env_int(self, temp_config_dir):
        """Test getting integer values from environment."""
        manager = ConfigManager(temp_config_dir)
        
        with patch.dict(os.environ, {'TEST_INT': '123'}):
            assert manager._get_env_int('TEST_INT', 456) == 123
        
        with patch.dict(os.environ, {'TEST_INT': 'invalid'}):
            assert manager._get_env_int('TEST_INT', 456) == 456
        
        with patch.dict(os.environ, {}, clear=True):
            assert manager._get_env_int('TEST_INT', 456) == 456
    
    def test_get_app_config(self, temp_config_dir):
        """Test getting application configuration."""
        manager = ConfigManager(temp_config_dir)
        app_config = manager.get_app_config()
        
        assert "development_mode" in app_config
        assert "log_level" in app_config
        assert "max_file_size_mb" in app_config
        assert "allowed_extensions" in app_config
        assert "flask_config" in app_config
        assert "google_cloud" in app_config
        assert "zoho_workdrive" in app_config
        
        assert isinstance(app_config["allowed_extensions"], list)
        assert isinstance(app_config["flask_config"], dict)
    
    def test_merge_configs(self, temp_config_dir):
        """Test merging configurations."""
        manager = ConfigManager(temp_config_dir)
        
        base_config = {
            "a": 1,
            "b": {"x": 1, "y": 2},
            "c": [1, 2, 3]
        }
        
        override_config = {
            "b": {"y": 20, "z": 30},
            "d": 4
        }
        
        merged = manager.merge_configs(base_config, override_config)
        
        assert merged["a"] == 1
        assert merged["b"]["x"] == 1  # From base
        assert merged["b"]["y"] == 20  # Overridden
        assert merged["b"]["z"] == 30  # Added
        assert merged["c"] == [1, 2, 3]  # From base
        assert merged["d"] == 4  # Added
    
    def test_validate_runtime_config_valid(self, temp_config_dir, sample_config):
        """Test runtime configuration validation with valid config."""
        manager = ConfigManager(temp_config_dir)
        # Should not raise exception
        manager.validate_runtime_config(sample_config)
    
    def test_validate_runtime_config_invalid_sheet_name(self, temp_config_dir):
        """Test runtime validation with invalid sheet name."""
        invalid_config = {
            "version": "1.0",
            "sheet_configs": {
                "": {  # Empty sheet name
                    "subtables": []
                }
            },
            "global_settings": {}
        }
        
        manager = ConfigManager(temp_config_dir)
        with pytest.raises(ConfigurationError):
            manager.validate_runtime_config(invalid_config)
    
    def test_validate_runtime_config_missing_subtables(self, temp_config_dir):
        """Test runtime validation with missing subtables."""
        invalid_config = {
            "version": "1.0",
            "sheet_configs": {
                "TestSheet": {
                    # Missing subtables
                }
            },
            "global_settings": {}
        }
        
        manager = ConfigManager(temp_config_dir)
        with pytest.raises(ConfigurationError):
            manager.validate_runtime_config(invalid_config)
    
    def test_validate_subtable_config_missing_fields(self, temp_config_dir):
        """Test subtable validation with missing required fields."""
        manager = ConfigManager(temp_config_dir)
        
        invalid_subtable = {
            "name": "test",
            "type": "key_value_pairs",
            # Missing header_search and data_extraction
        }
        
        with pytest.raises(ConfigurationError):
            manager._validate_subtable_config(invalid_subtable)
    
    def test_validate_subtable_config_invalid_type(self, temp_config_dir):
        """Test subtable validation with invalid type."""
        manager = ConfigManager(temp_config_dir)
        
        invalid_subtable = {
            "name": "test",
            "type": "invalid_type",
            "header_search": {"method": "contains_text"},
            "data_extraction": {}
        }
        
        with pytest.raises(ConfigurationError):
            manager._validate_subtable_config(invalid_subtable)
    
    def test_validate_subtable_config_invalid_search_method(self, temp_config_dir):
        """Test subtable validation with invalid search method."""
        manager = ConfigManager(temp_config_dir)
        
        invalid_subtable = {
            "name": "test",
            "type": "key_value_pairs",
            "header_search": {"method": "invalid_method"},
            "data_extraction": {}
        }
        
        with pytest.raises(ConfigurationError):
            manager._validate_subtable_config(invalid_subtable)
    
    def test_validate_global_settings_negative_delay(self, temp_config_dir):
        """Test global settings validation with negative delay."""
        manager = ConfigManager(temp_config_dir)
        
        invalid_settings = {
            "temp_file_cleanup": {
                "delay_seconds": -1
            }
        }
        
        with pytest.raises(ConfigurationError):
            manager._validate_global_settings(invalid_settings)
    
    def test_validate_global_settings_invalid_image_size(self, temp_config_dir):
        """Test global settings validation with invalid image size."""
        manager = ConfigManager(temp_config_dir)
        
        invalid_settings = {
            "image_extraction": {
                "max_size_mb": 0
            }
        }
        
        with pytest.raises(ConfigurationError):
            manager._validate_global_settings(invalid_settings)
    
    def test_clear_cache(self, temp_config_dir, sample_config):
        """Test clearing configuration cache."""
        manager = ConfigManager(temp_config_dir)
        
        # Load config to populate cache
        manager._config_cache["test"] = sample_config
        assert len(manager._config_cache) > 0
        
        # Clear cache
        manager.clear_cache()
        assert len(manager._config_cache) == 0
    
    def test_get_cached_configs(self, temp_config_dir, sample_config):
        """Test getting list of cached configurations."""
        manager = ConfigManager(temp_config_dir)
        
        # Add configs to cache
        manager._config_cache["config1"] = sample_config
        manager._config_cache["config2"] = sample_config
        
        cached_configs = manager.get_cached_configs()
        assert "config1" in cached_configs
        assert "config2" in cached_configs
        assert len(cached_configs) == 2
    
    @patch.dict(os.environ, {}, clear=True)
    def test_load_environment_no_files(self, temp_config_dir):
        """Test loading environment when no .env files exist."""
        manager = ConfigManager(temp_config_dir)
        # Should not raise exception even if no .env files exist
        manager._load_environment()
    
    @patch('src.config_manager.load_dotenv')
    def test_load_environment_with_files(self, mock_load_dotenv, temp_config_dir):
        """Test loading environment with .env files."""
        # Create .env file
        env_file = os.path.join(os.getcwd(), '.env')
        with open(env_file, 'w') as f:
            f.write('TEST_VAR=test_value\n')
        
        try:
            manager = ConfigManager(temp_config_dir)
            manager._load_environment()
            
            # Should have called load_dotenv
            assert mock_load_dotenv.called
        finally:
            if os.path.exists(env_file):
                os.unlink(env_file)