"""Tests for Excel processor module."""

import os
import pytest
import tempfile
from unittest.mock import Mock, patch
import pandas as pd
from openpyxl import Workbook

from src.excel_processor import ExcelProcessor
from src.utils.exceptions import ExcelProcessingError, ValidationError


class TestExcelProcessor:
    """Test cases for ExcelProcessor class."""
    
    @pytest.fixture
    def sample_excel_file(self):
        """Create a sample Excel file for testing."""
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_file:
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "TestSheet"
            
            # Add test data
            worksheet['A1'] = 'Client'
            worksheet['B1'] = 'Type'
            worksheet['C1'] = 'Date'
            worksheet['A2'] = 'Test Client'
            worksheet['B2'] = 'Premium'
            worksheet['C2'] = '2023-01-01'
            
            # Add table data
            worksheet['A5'] = 'Data'
            worksheet['A6'] = 'ID'
            worksheet['B6'] = 'Name'
            worksheet['C6'] = 'Value'
            worksheet['A7'] = '1'
            worksheet['B7'] = 'Item 1'
            worksheet['C7'] = '100'
            worksheet['A8'] = '2'
            worksheet['B8'] = 'Item 2'
            worksheet['C8'] = '200'
            
            workbook.save(tmp_file.name)
            workbook.close()
            
            yield tmp_file.name
            
            # Cleanup
            if os.path.exists(tmp_file.name):
                os.unlink(tmp_file.name)
    
    @pytest.fixture
    def test_config(self):
        """Test configuration for Excel processing."""
        return {
            "TestSheet": {
                "subtables": [
                    {
                        "name": "test_client_info",
                        "type": "key_value_pairs",
                        "header_search": {
                            "method": "contains_text",
                            "text": "Client",
                            "column": "A",
                            "search_range": "A1:A5"
                        },
                        "data_extraction": {
                            "orientation": "horizontal",
                            "headers_row_offset": 0,
                            "data_row_offset": 1,
                            "max_columns": 3,
                            "column_mappings": {
                                "Client": "client_name",
                                "Type": "client_type",
                                "Date": "created_date"
                            }
                        }
                    },
                    {
                        "name": "test_data_table",
                        "type": "table",
                        "header_search": {
                            "method": "contains_text",
                            "text": "Data",
                            "column": "A",
                            "search_range": "A5:A10"
                        },
                        "data_extraction": {
                            "orientation": "vertical",
                            "headers_row_offset": 1,
                            "data_row_offset": 2,
                            "max_columns": 3,
                            "max_rows": 5,
                            "column_mappings": {
                                "ID": "id",
                                "Name": "name",
                                "Value": "value"
                            }
                        }
                    }
                ]
            }
        }
    
    def test_init_valid_file(self, sample_excel_file):
        """Test Excel processor initialization with valid file."""
        processor = ExcelProcessor(sample_excel_file)
        assert processor.file_path == sample_excel_file
        assert processor.workbook is not None
        processor.close()
    
    def test_init_invalid_file(self):
        """Test Excel processor initialization with invalid file."""
        with pytest.raises(ExcelProcessingError):
            ExcelProcessor("nonexistent_file.xlsx")
    
    def test_get_sheet_names(self, sample_excel_file):
        """Test getting sheet names from workbook."""
        processor = ExcelProcessor(sample_excel_file)
        sheet_names = processor.get_sheet_names()
        assert "TestSheet" in sheet_names
        processor.close()
    
    def test_extract_data(self, sample_excel_file, test_config):
        """Test data extraction with configuration."""
        processor = ExcelProcessor(sample_excel_file)
        try:
            extracted_data = processor.extract_data(test_config)
            
            # Verify structure
            assert "TestSheet" in extracted_data
            sheet_data = extracted_data["TestSheet"]
            
            # Verify client info
            assert "test_client_info" in sheet_data
            client_info = sheet_data["test_client_info"]
            assert client_info["client_name"] == "Test Client"
            assert client_info["client_type"] == "Premium"
            assert client_info["created_date"] == "2023-01-01"
            
            # Verify table data
            assert "test_data_table" in sheet_data
            table_data = sheet_data["test_data_table"]
            assert len(table_data) == 2
            assert table_data[0]["id"] == 1
            assert table_data[0]["name"] == "Item 1"
            assert table_data[0]["value"] == 100
            
        finally:
            processor.close()
    
    def test_find_header_location_contains_text(self, sample_excel_file):
        """Test finding header location using contains_text method."""
        processor = ExcelProcessor(sample_excel_file)
        try:
            worksheet = processor.workbook["TestSheet"]
            search_config = {
                "method": "contains_text",
                "text": "Client",
                "search_range": "A1:A5"
            }
            
            location = processor._find_header_location(worksheet, search_config)
            assert location == (1, 1)  # Row 1, Column 1 (A1)
            
        finally:
            processor.close()
    
    def test_find_header_location_exact_match(self, sample_excel_file):
        """Test finding header location using exact_match method."""
        processor = ExcelProcessor(sample_excel_file)
        try:
            worksheet = processor.workbook["TestSheet"]
            search_config = {
                "method": "exact_match",
                "text": "Client",
                "search_range": "A1:A5"
            }
            
            location = processor._find_header_location(worksheet, search_config)
            assert location == (1, 1)  # Row 1, Column 1 (A1)
            
        finally:
            processor.close()
    
    def test_extract_key_value_pairs_horizontal(self, sample_excel_file):
        """Test extracting key-value pairs in horizontal orientation."""
        processor = ExcelProcessor(sample_excel_file)
        try:
            worksheet = processor.workbook["TestSheet"]
            header_location = (1, 1)  # A1
            config = {
                "orientation": "horizontal",
                "headers_row_offset": 0,
                "data_row_offset": 1,
                "max_columns": 3,
                "column_mappings": {
                    "Client": "client_name",
                    "Type": "client_type",
                    "Date": "created_date"
                }
            }
            
            data = processor._extract_key_value_pairs(worksheet, header_location, config)
            assert data["client_name"] == "Test Client"
            assert data["client_type"] == "Premium"
            assert data["created_date"] == "2023-01-01"
            
        finally:
            processor.close()
    
    def test_extract_table_data(self, sample_excel_file):
        """Test extracting table data."""
        processor = ExcelProcessor(sample_excel_file)
        try:
            worksheet = processor.workbook["TestSheet"]
            header_location = (6, 1)  # A6 (headers row)
            config = {
                "headers_row_offset": 0,
                "data_row_offset": 1,
                "max_columns": 3,
                "max_rows": 5,
                "column_mappings": {
                    "ID": "id",
                    "Name": "name",
                    "Value": "value"
                }
            }
            
            data = processor._extract_table_data(worksheet, header_location, config)
            assert len(data) == 2
            assert data[0]["id"] == 1
            assert data[0]["name"] == "Item 1"
            assert data[0]["value"] == 100
            assert data[1]["id"] == 2
            assert data[1]["name"] == "Item 2"
            assert data[1]["value"] == 200
            
        finally:
            processor.close()
    
    def test_get_cell_value(self, sample_excel_file):
        """Test getting value from specific cell."""
        processor = ExcelProcessor(sample_excel_file)
        try:
            value = processor.get_cell_value("TestSheet", "A1")
            assert value == "Client"
            
            value = processor.get_cell_value("TestSheet", "A2")
            assert value == "Test Client"
            
        finally:
            processor.close()
    
    def test_get_cell_value_invalid_sheet(self, sample_excel_file):
        """Test getting cell value from invalid sheet."""
        processor = ExcelProcessor(sample_excel_file)
        try:
            with pytest.raises(ValidationError):
                processor.get_cell_value("NonexistentSheet", "A1")
        finally:
            processor.close()
    
    def test_get_range_values(self, sample_excel_file):
        """Test getting values from cell range."""
        processor = ExcelProcessor(sample_excel_file)
        try:
            values = processor.get_range_values("TestSheet", "A1:C2")
            assert len(values) == 2  # 2 rows
            assert len(values[0]) == 3  # 3 columns
            assert values[0][0] == "Client"
            assert values[0][1] == "Type"
            assert values[0][2] == "Date"
            assert values[1][0] == "Test Client"
            assert values[1][1] == "Premium"
            assert values[1][2] == "2023-01-01"
            
        finally:
            processor.close()
    
    def test_to_dataframe(self, sample_excel_file):
        """Test converting Excel sheet to pandas DataFrame."""
        processor = ExcelProcessor(sample_excel_file)
        try:
            df = processor.to_dataframe("TestSheet", header_row=0)
            assert isinstance(df, pd.DataFrame)
            assert "Client" in df.columns
            assert "Type" in df.columns
            assert "Date" in df.columns
            assert len(df) > 0
            
        finally:
            processor.close()
    
    def test_extract_images_no_images(self, sample_excel_file):
        """Test image extraction when no images are present."""
        processor = ExcelProcessor(sample_excel_file)
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                images = processor.extract_images(temp_dir)
                assert isinstance(images, dict)
                # Should be empty since no images in test file
                
        finally:
            processor.close()
    
    def test_invalid_config_missing_subtables(self, sample_excel_file):
        """Test with invalid configuration missing subtables."""
        processor = ExcelProcessor(sample_excel_file)
        try:
            invalid_config = {
                "TestSheet": {
                    # Missing subtables
                }
            }
            
            with pytest.raises(ValidationError):
                processor.extract_data(invalid_config)
                
        finally:
            processor.close()
    
    def test_invalid_cell_range(self, sample_excel_file):
        """Test with invalid cell range."""
        processor = ExcelProcessor(sample_excel_file)
        try:
            with pytest.raises(ValidationError):
                processor.get_range_values("TestSheet", "INVALID_RANGE")
                
        finally:
            processor.close()
    
    def test_nonexistent_sheet_in_config(self, sample_excel_file):
        """Test with configuration referencing non-existent sheet."""
        processor = ExcelProcessor(sample_excel_file)
        try:
            config = {
                "NonexistentSheet": {
                    "subtables": [
                        {
                            "name": "test",
                            "type": "key_value_pairs",
                            "header_search": {"method": "contains_text", "text": "test"},
                            "data_extraction": {"orientation": "horizontal"}
                        }
                    ]
                }
            }
            
            # Should not raise exception, just skip missing sheet
            extracted_data = processor.extract_data(config)
            assert "NonexistentSheet" not in extracted_data or extracted_data["NonexistentSheet"] == {}
            
        finally:
            processor.close()