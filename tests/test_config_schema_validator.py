"""Unit tests for configuration schema validator."""

import pytest
from src.config_schema_validator import ConfigSchemaValidator, validate_config_file, create_range_image_schema
from src.utils.exceptions import ValidationError


class TestConfigSchemaValidator:
    """Test cases for ConfigSchemaValidator."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.validator = ConfigSchemaValidator()
    
    def test_validate_basic_structure_success(self):
        """Test successful basic structure validation."""
        config = {
            "version": "1.1",
            "sheet_configs": {
                "Sheet1": {"subtables": []}
            }
        }
        
        is_valid = self.validator.validate_config(config)
        assert is_valid is True
        assert len(self.validator.get_validation_errors()) == 0
    
    def test_validate_basic_structure_missing_required(self):
        """Test basic structure validation with missing required fields."""
        config = {
            "version": "1.1"
            # Missing sheet_configs
        }
        
        is_valid = self.validator.validate_config(config)
        assert is_valid is False
        errors = self.validator.get_validation_errors()
        assert any("Missing required field: sheet_configs" in error for error in errors)
    
    def test_validate_basic_structure_invalid_version(self):
        """Test basic structure validation with invalid version."""
        config = {
            "version": "2.0",  # Unsupported version
            "sheet_configs": {}
        }
        
        is_valid = self.validator.validate_config(config)
        assert is_valid is False
        errors = self.validator.get_validation_errors()
        assert any("Unsupported version" in error for error in errors)
    
    def test_validate_range_images_success(self):
        """Test successful range images validation."""
        config = {
            "version": "1.1",
            "sheet_configs": {"Sheet1": {}},
            "range_images": [
                {
                    "field_name": "test_field",
                    "sheet_name": "Sheet1",
                    "range": "A1:C10",
                    "output_format": "png",
                    "dpi": 150
                }
            ]
        }
        
        is_valid = self.validator.validate_config(config)
        assert is_valid is True
    
    def test_validate_range_images_missing_required(self):
        """Test range images validation with missing required fields."""
        config = {
            "version": "1.1",
            "sheet_configs": {"Sheet1": {}},
            "range_images": [
                {
                    "field_name": "test_field",
                    # Missing sheet_name and range
                    "output_format": "png"
                }
            ]
        }
        
        is_valid = self.validator.validate_config(config)
        assert is_valid is False
        errors = self.validator.get_validation_errors()
        assert any("Missing required field 'sheet_name'" in error for error in errors)
        assert any("Missing required field 'range'" in error for error in errors)
    
    def test_validate_range_images_duplicate_field_names(self):
        """Test range images validation with duplicate field names."""
        config = {
            "version": "1.1",
            "sheet_configs": {"Sheet1": {}},
            "range_images": [
                {
                    "field_name": "duplicate_field",
                    "sheet_name": "Sheet1",
                    "range": "A1:C10"
                },
                {
                    "field_name": "duplicate_field",  # Duplicate
                    "sheet_name": "Sheet1",
                    "range": "D1:F10"
                }
            ]
        }
        
        is_valid = self.validator.validate_config(config)
        assert is_valid is False
        errors = self.validator.get_validation_errors()
        assert any("Duplicate field_name" in error for error in errors)
    
    def test_validate_range_config_fields_invalid_range(self):
        """Test range config field validation with invalid range."""
        config = {
            "version": "1.1",
            "sheet_configs": {"Sheet1": {}},
            "range_images": [
                {
                    "field_name": "test_field",
                    "sheet_name": "Sheet1",
                    "range": "InvalidRange"
                }
            ]
        }
        
        is_valid = self.validator.validate_config(config)
        assert is_valid is False
        errors = self.validator.get_validation_errors()
        assert any("Invalid range format" in error for error in errors)
    
    def test_validate_range_config_fields_invalid_format(self):
        """Test range config field validation with invalid output format."""
        config = {
            "version": "1.1",
            "sheet_configs": {"Sheet1": {}},
            "range_images": [
                {
                    "field_name": "test_field",
                    "sheet_name": "Sheet1",
                    "range": "A1:C10",
                    "output_format": "invalid_format"
                }
            ]
        }
        
        is_valid = self.validator.validate_config(config)
        assert is_valid is False
        errors = self.validator.get_validation_errors()
        assert any("Invalid output_format" in error for error in errors)
    
    def test_validate_range_config_fields_invalid_dpi(self):
        """Test range config field validation with invalid DPI."""
        config = {
            "version": "1.1",
            "sheet_configs": {"Sheet1": {}},
            "range_images": [
                {
                    "field_name": "test_field",
                    "sheet_name": "Sheet1",
                    "range": "A1:C10",
                    "dpi": 1000  # Too high
                }
            ]
        }
        
        is_valid = self.validator.validate_config(config)
        assert is_valid is False
        errors = self.validator.get_validation_errors()
        assert any("DPI must be integer between" in error for error in errors)
    
    def test_validate_range_config_fields_invalid_dimensions(self):
        """Test range config field validation with invalid dimensions."""
        config = {
            "version": "1.1",
            "sheet_configs": {"Sheet1": {}},
            "range_images": [
                {
                    "field_name": "test_field",
                    "sheet_name": "Sheet1",
                    "range": "A1:C10",
                    "width": -100  # Invalid negative width
                }
            ]
        }
        
        is_valid = self.validator.validate_config(config)
        assert is_valid is False
        errors = self.validator.get_validation_errors()
        assert any("width must be positive integer" in error for error in errors)
    
    def test_validate_global_settings_range_images(self):
        """Test global settings validation for range images."""
        config = {
            "version": "1.1",
            "sheet_configs": {"Sheet1": {}},
            "global_settings": {
                "range_images": {
                    "enabled": True,
                    "max_range_cells": 5000,
                    "default_dpi": 200,
                    "default_format": "jpg"
                }
            }
        }
        
        is_valid = self.validator.validate_config(config)
        assert is_valid is True
    
    def test_validate_global_settings_invalid_dpi(self):
        """Test global settings validation with invalid default DPI."""
        config = {
            "version": "1.1",
            "sheet_configs": {"Sheet1": {}},
            "global_settings": {
                "range_images": {
                    "default_dpi": 1000  # Too high
                }
            }
        }
        
        is_valid = self.validator.validate_config(config)
        assert is_valid is False
        errors = self.validator.get_validation_errors()
        assert any("default_dpi must be between" in error for error in errors)
    
    def test_cross_validate_range_images_sheet_not_found(self):
        """Test cross-validation when range image references non-existent sheet."""
        config = {
            "version": "1.1",
            "sheet_configs": {"Sheet1": {}},  # Only Sheet1 exists
            "range_images": [
                {
                    "field_name": "test_field",
                    "sheet_name": "NonExistentSheet",  # References non-existent sheet
                    "range": "A1:C10"
                }
            ]
        }
        
        is_valid = self.validator.validate_config(config)
        assert is_valid is False
        errors = self.validator.get_validation_errors()
        assert any("Sheet 'NonExistentSheet' not found in sheet_configs" in error for error in errors)
    
    def test_is_valid_excel_range_valid_formats(self):
        """Test valid Excel range format detection."""
        assert self.validator._is_valid_excel_range("A1:C10") is True
        assert self.validator._is_valid_excel_range("AA1:ZZ100") is True
        assert self.validator._is_valid_excel_range("a1:c10") is True  # Lowercase
    
    def test_is_valid_excel_range_invalid_formats(self):
        """Test invalid Excel range format detection."""
        assert self.validator._is_valid_excel_range("InvalidRange") is False
        assert self.validator._is_valid_excel_range("A1-C10") is False
        assert self.validator._is_valid_excel_range("A1:") is False
        assert self.validator._is_valid_excel_range(":C10") is False
        assert self.validator._is_valid_excel_range("1A:10C") is False
    
    def test_get_range_cell_count_valid_ranges(self):
        """Test cell count calculation for valid ranges."""
        assert self.validator._get_range_cell_count("A1:C3") == 9  # 3x3
        assert self.validator._get_range_cell_count("A1:A10") == 10  # 1x10
        assert self.validator._get_range_cell_count("A1:J1") == 10  # 10x1
        assert self.validator._get_range_cell_count("B2:D4") == 9  # 3x3
    
    def test_get_range_cell_count_invalid_ranges(self):
        """Test cell count calculation for invalid ranges."""
        assert self.validator._get_range_cell_count("InvalidRange") == 0
        assert self.validator._get_range_cell_count("A1-C3") == 0
        assert self.validator._get_range_cell_count("") == 0
    
    def test_column_letters_to_number(self):
        """Test column letter to number conversion."""
        assert self.validator._column_letters_to_number("A") == 1
        assert self.validator._column_letters_to_number("B") == 2
        assert self.validator._column_letters_to_number("Z") == 26
        assert self.validator._column_letters_to_number("AA") == 27
        assert self.validator._column_letters_to_number("AB") == 28


class TestHelperFunctions:
    """Test cases for helper functions."""
    
    def test_validate_config_file_success(self):
        """Test successful config file validation."""
        config = {
            "version": "1.1",
            "sheet_configs": {"Sheet1": {}},
            "range_images": [
                {
                    "field_name": "test_field",
                    "sheet_name": "Sheet1",
                    "range": "A1:C10"
                }
            ]
        }
        
        is_valid, errors = validate_config_file(config)
        assert is_valid is True
        assert len(errors) == 0
    
    def test_validate_config_file_with_errors(self):
        """Test config file validation with errors."""
        config = {
            "version": "1.1",
            "sheet_configs": {},
            "range_images": [
                {
                    "field_name": "",  # Invalid empty field name
                    "sheet_name": "Sheet1",
                    "range": "InvalidRange"  # Invalid range format
                }
            ]
        }
        
        is_valid, errors = validate_config_file(config)
        assert is_valid is False
        assert len(errors) > 0
    
    def test_create_range_image_schema_success(self):
        """Test successful range image schema creation."""
        config_dict = {
            "field_name": "test_field",
            "sheet_name": "Sheet1",
            "range": "A1:C10",
            "include_headers": False,
            "output_format": "jpg",
            "dpi": 300,
            "width": 800,
            "height": 600
        }
        
        schema = create_range_image_schema(config_dict)
        assert schema.field_name == "test_field"
        assert schema.sheet_name == "Sheet1"
        assert schema.range == "A1:C10"
        assert schema.include_headers is False
        assert schema.output_format == "jpg"
        assert schema.dpi == 300
        assert schema.width == 800
        assert schema.height == 600
    
    def test_create_range_image_schema_with_defaults(self):
        """Test range image schema creation with default values."""
        config_dict = {
            "field_name": "test_field",
            "sheet_name": "Sheet1",
            "range": "A1:C10"
            # Other fields will use defaults
        }
        
        schema = create_range_image_schema(config_dict)
        assert schema.include_headers is True  # Default
        assert schema.output_format == "png"  # Default
        assert schema.dpi == 150  # Default
        assert schema.fit_to_content is True  # Default
        assert schema.width is None  # Default
        assert schema.height is None  # Default
    
    def test_create_range_image_schema_missing_required(self):
        """Test range image schema creation with missing required field."""
        config_dict = {
            "field_name": "test_field",
            # Missing sheet_name and range
        }
        
        with pytest.raises(ValidationError, match="Missing required field"):
            create_range_image_schema(config_dict)
    
    def test_create_range_image_schema_invalid_data(self):
        """Test range image schema creation with invalid data."""
        config_dict = {
            "field_name": "test_field",
            "sheet_name": "Sheet1",
            "range": "A1:C10",
            "invalid_field": "invalid_value"  # This should be ignored
        }
        
        # Should succeed but ignore invalid field
        schema = create_range_image_schema(config_dict)
        assert schema.field_name == "test_field"
        assert not hasattr(schema, 'invalid_field')