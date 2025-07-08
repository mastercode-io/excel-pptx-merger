"""Unit tests for Excel range exporter."""

import pytest
from unittest.mock import Mock, patch, MagicMock
import tempfile
import os

from src.excel_range_exporter import (
    ExcelRangeExporter, 
    RangeImageConfig, 
    RangeImageResult,
    create_range_configs_from_dict,
    validate_range_configs
)
from src.utils.exceptions import ValidationError


class TestRangeImageConfig:
    """Test cases for RangeImageConfig dataclass."""
    
    def test_required_fields(self):
        """Test RangeImageConfig with required fields only."""
        config = RangeImageConfig(
            field_name="test_field",
            sheet_name="Sheet1", 
            range="A1:C10"
        )
        
        assert config.field_name == "test_field"
        assert config.sheet_name == "Sheet1"
        assert config.range == "A1:C10"
        # Test default values
        assert config.include_headers is True
        assert config.output_format == "png"
        assert config.dpi == 150
        assert config.fit_to_content is True
        assert config.width is None
        assert config.height is None
    
    def test_all_fields(self):
        """Test RangeImageConfig with all fields specified."""
        config = RangeImageConfig(
            field_name="test_field",
            sheet_name="Sheet1",
            range="A1:C10",
            include_headers=False,
            output_format="jpg",
            dpi=300,
            fit_to_content=False,
            width=800,
            height=600
        )
        
        assert config.include_headers is False
        assert config.output_format == "jpg"
        assert config.dpi == 300
        assert config.fit_to_content is False
        assert config.width == 800
        assert config.height == 600


class TestRangeImageResult:
    """Test cases for RangeImageResult dataclass."""
    
    def test_success_result(self):
        """Test RangeImageResult for successful export."""
        result = RangeImageResult(
            field_name="test_field",
            image_path="/path/to/image.png",
            image_data=b"fake_image_data",
            width=800,
            height=600,
            range_dimensions=(10, 3),
            success=True
        )
        
        assert result.success is True
        assert result.error_message is None
        assert result.range_dimensions == (10, 3)
    
    def test_error_result(self):
        """Test RangeImageResult for failed export."""
        result = RangeImageResult(
            field_name="test_field",
            image_path="",
            image_data=b"",
            width=0,
            height=0,
            range_dimensions=(0, 0),
            success=False,
            error_message="Test error"
        )
        
        assert result.success is False
        assert result.error_message == "Test error"


class TestExcelRangeExporter:
    """Test cases for ExcelRangeExporter."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.exporter = ExcelRangeExporter(
            client_id="test_client",
            client_secret="test_secret",
            tenant_id="test_tenant"
        )
    
    def test_init(self):
        """Test exporter initialization."""
        assert self.exporter.graph_client is not None
        assert self.exporter.temp_manager is not None
    
    @patch('os.path.exists', return_value=False)
    def test_export_ranges_missing_workbook(self, mock_exists):
        """Test export with missing workbook file."""
        config = RangeImageConfig("test", "Sheet1", "A1:C10")
        
        with pytest.raises(Exception, match="Workbook not found"):
            self.exporter.export_ranges("/missing/file.xlsx", [config])
    
    @patch('os.path.exists', return_value=True)
    def test_export_ranges_empty_configs(self, mock_exists):
        """Test export with empty configuration list."""
        results = self.exporter.export_ranges("/path/to/file.xlsx", [])
        assert results == []
    
    @patch('os.path.exists', return_value=True)
    def test_export_ranges_success(self, mock_exists):
        """Test successful range export."""
        config = RangeImageConfig("test_field", "Sheet1", "A1:C10")
        
        # Mock graph client methods
        mock_graph_client = Mock()
        mock_graph_client.upload_workbook_to_onedrive.return_value = "test_item_id"
        mock_graph_client.get_worksheet_names.return_value = ["Sheet1", "Sheet2"]
        mock_graph_client.validate_range.return_value = True
        mock_graph_client.get_range_dimensions.return_value = (10, 3)
        mock_graph_client.render_range_as_image.return_value = b"fake_image_data"
        mock_graph_client.cleanup_temp_file.return_value = None
        
        self.exporter.graph_client = mock_graph_client
        
        # Mock temp file creation
        with patch.object(self.exporter, '_save_image_to_temp_file', return_value="/tmp/test_image.png"):
            with patch.object(self.exporter, '_get_image_dimensions', return_value=(800, 600)):
                results = self.exporter.export_ranges("/path/to/file.xlsx", [config])
                
                assert len(results) == 1
                result = results[0]
                assert result.success is True
                assert result.field_name == "test_field"
                assert result.image_path == "/tmp/test_image.png"
                assert result.width == 800
                assert result.height == 600
                assert result.range_dimensions == (10, 3)
    
    @patch('os.path.exists', return_value=True)
    def test_export_ranges_sheet_not_found(self, mock_exists):
        """Test export with non-existent sheet."""
        config = RangeImageConfig("test_field", "MissingSheet", "A1:C10")
        
        # Mock graph client methods
        mock_graph_client = Mock()
        mock_graph_client.upload_workbook_to_onedrive.return_value = "test_item_id"
        mock_graph_client.get_worksheet_names.return_value = ["Sheet1", "Sheet2"]  # MissingSheet not in list
        mock_graph_client.cleanup_temp_file.return_value = None
        
        self.exporter.graph_client = mock_graph_client
        
        results = self.exporter.export_ranges("/path/to/file.xlsx", [config])
        
        assert len(results) == 1
        result = results[0]
        assert result.success is False
        assert "not found" in result.error_message
    
    @patch('os.path.exists', return_value=True)
    def test_export_ranges_invalid_range(self, mock_exists):
        """Test export with invalid range."""
        config = RangeImageConfig("test_field", "Sheet1", "InvalidRange")
        
        # Mock graph client methods
        mock_graph_client = Mock()
        mock_graph_client.upload_workbook_to_onedrive.return_value = "test_item_id"
        mock_graph_client.get_worksheet_names.return_value = ["Sheet1"]
        mock_graph_client.validate_range.return_value = False  # Invalid range
        mock_graph_client.cleanup_temp_file.return_value = None
        
        self.exporter.graph_client = mock_graph_client
        
        results = self.exporter.export_ranges("/path/to/file.xlsx", [config])
        
        assert len(results) == 1
        result = results[0]
        assert result.success is False
        assert "invalid or empty" in result.error_message
    
    def test_save_image_to_temp_file(self):
        """Test saving image data to temporary file."""
        image_data = b"fake_image_data"
        
        with patch.object(self.exporter.temp_manager, 'create_temp_file', return_value="/tmp/test.png"):
            with patch('builtins.open', create=True) as mock_open:
                mock_file = Mock()
                mock_open.return_value.__enter__.return_value = mock_file
                
                result_path = self.exporter._save_image_to_temp_file(image_data, "test_field", "png")
                
                assert result_path == "/tmp/test.png"
                mock_file.write.assert_called_once_with(image_data)
    
    def test_get_image_dimensions_success(self):
        """Test getting image dimensions from binary data."""
        # Mock PIL Image
        with patch('PIL.Image.open') as mock_image_open:
            mock_image = Mock()
            mock_image.size = (800, 600)
            mock_image_open.return_value = mock_image
            
            width, height = self.exporter._get_image_dimensions(b"fake_image_data")
            
            assert width == 800
            assert height == 600
    
    def test_get_image_dimensions_error(self):
        """Test getting image dimensions with error."""
        with patch('PIL.Image.open', side_effect=Exception("Invalid image")):
            width, height = self.exporter._get_image_dimensions(b"invalid_data")
            
            assert width == 0
            assert height == 0
    
    def test_validate_config_success(self):
        """Test successful config validation."""
        config = RangeImageConfig("test_field", "Sheet1", "A1:C10")
        
        is_valid = self.exporter.validate_config(config)
        assert is_valid is True
    
    def test_validate_config_missing_field_name(self):
        """Test config validation with missing field name."""
        config = RangeImageConfig("", "Sheet1", "A1:C10")
        
        with pytest.raises(ValidationError, match="field_name is required"):
            self.exporter.validate_config(config)
    
    def test_validate_config_invalid_range(self):
        """Test config validation with invalid range format."""
        config = RangeImageConfig("test", "Sheet1", "InvalidRange")
        
        with pytest.raises(ValidationError, match="Invalid range format"):
            self.exporter.validate_config(config)
    
    def test_validate_config_invalid_format(self):
        """Test config validation with invalid output format."""
        config = RangeImageConfig("test", "Sheet1", "A1:C10", output_format="invalid")
        
        with pytest.raises(ValidationError, match="Invalid output format"):
            self.exporter.validate_config(config)
    
    def test_validate_config_invalid_dpi(self):
        """Test config validation with invalid DPI."""
        config = RangeImageConfig("test", "Sheet1", "A1:C10", dpi=1000)
        
        with pytest.raises(ValidationError, match="DPI must be between"):
            self.exporter.validate_config(config)
    
    def test_is_valid_range_format_valid(self):
        """Test valid range format detection."""
        assert self.exporter._is_valid_range_format("A1:C10") is True
        assert self.exporter._is_valid_range_format("AA1:ZZ100") is True
        assert self.exporter._is_valid_range_format("a1:c10") is True  # Should handle lowercase
    
    def test_is_valid_range_format_invalid(self):
        """Test invalid range format detection."""
        assert self.exporter._is_valid_range_format("InvalidRange") is False
        assert self.exporter._is_valid_range_format("A1-C10") is False
        assert self.exporter._is_valid_range_format("A1:") is False
        assert self.exporter._is_valid_range_format(":C10") is False
    
    def test_cleanup_temp_files(self):
        """Test temporary file cleanup."""
        with patch.object(self.exporter.temp_manager, 'cleanup_all') as mock_cleanup:
            self.exporter.cleanup_temp_files()
            mock_cleanup.assert_called_once()


class TestHelperFunctions:
    """Test cases for helper functions."""
    
    def test_create_range_configs_from_dict_success(self):
        """Test creating range configs from dictionary data."""
        configs_data = [
            {
                "field_name": "test1",
                "sheet_name": "Sheet1",
                "range": "A1:C10"
            },
            {
                "field_name": "test2",
                "sheet_name": "Sheet2", 
                "range": "B2:D20",
                "dpi": 300
            }
        ]
        
        configs = create_range_configs_from_dict(configs_data)
        
        assert len(configs) == 2
        assert configs[0].field_name == "test1"
        assert configs[0].dpi == 150  # Default
        assert configs[1].field_name == "test2"
        assert configs[1].dpi == 300  # Specified
    
    def test_create_range_configs_from_dict_invalid(self):
        """Test creating range configs with invalid data."""
        configs_data = [
            {
                "field_name": "test1",
                # missing sheet_name and range
            }
        ]
        
        with pytest.raises(ValidationError):
            create_range_configs_from_dict(configs_data)
    
    def test_validate_range_configs_success(self):
        """Test successful validation of multiple configs."""
        configs = [
            RangeImageConfig("test1", "Sheet1", "A1:C10"),
            RangeImageConfig("test2", "Sheet2", "B2:D20")
        ]
        
        errors = validate_range_configs(configs)
        assert len(errors) == 0
    
    def test_validate_range_configs_duplicate_field_names(self):
        """Test validation with duplicate field names."""
        configs = [
            RangeImageConfig("test1", "Sheet1", "A1:C10"),
            RangeImageConfig("test1", "Sheet2", "B2:D20")  # Duplicate field_name
        ]
        
        errors = validate_range_configs(configs)
        assert len(errors) == 1
        assert "Duplicate field_name" in errors[0]
    
    def test_validate_range_configs_invalid_config(self):
        """Test validation with invalid individual config."""
        configs = [
            RangeImageConfig("", "Sheet1", "A1:C10"),  # Empty field_name
        ]
        
        errors = validate_range_configs(configs)
        assert len(errors) == 1
        assert "field_name is required" in errors[0]