"""Excel data extraction and processing module."""

import logging
import os
import re
from typing import Any, Dict, List, Optional, Tuple, Union
import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
import io

from .utils.exceptions import ExcelProcessingError, ValidationError
from .utils.validation import normalize_column_name, validate_cell_range, is_empty_cell_value

logger = logging.getLogger(__name__)


class ExcelProcessor:
    """Processes Excel files and extracts data according to configuration."""
    
    def __init__(self, file_path: str) -> None:
        """Initialize Excel processor with file path."""
        self.file_path = file_path
        self.workbook = None
        self.data_frame = None
        self._validate_file()
        
    def _validate_file(self) -> None:
        """Validate Excel file exists and is readable."""
        if not os.path.exists(self.file_path):
            raise ExcelProcessingError(f"Excel file not found: {self.file_path}")
        
        try:
            # Try to load workbook to validate format
            self.workbook = load_workbook(self.file_path, data_only=True)
            logger.info(f"Successfully loaded Excel file: {self.file_path}")
        except Exception as e:
            raise ExcelProcessingError(f"Invalid Excel file format: {e}")
    
    def get_sheet_names(self) -> List[str]:
        """Get list of sheet names in the workbook."""
        if not self.workbook:
            self._validate_file()
        return list(self.workbook.sheetnames)
    
    def extract_data(self, sheet_config: Dict[str, Any]) -> Dict[str, Any]:
        """Extract data from Excel sheet according to configuration."""
        try:
            extracted_data = {}
            
            for sheet_name, config in sheet_config.items():
                logger.info(f"Processing sheet: {sheet_name}")
                
                if sheet_name not in self.workbook.sheetnames:
                    logger.warning(f"Sheet '{sheet_name}' not found in workbook")
                    continue
                
                worksheet = self.workbook[sheet_name]
                sheet_data = self._process_sheet(worksheet, config)
                extracted_data[sheet_name] = sheet_data
            
            return extracted_data
            
        except Exception as e:
            raise ExcelProcessingError(f"Failed to extract data: {e}")
    
    def _process_sheet(self, worksheet: Worksheet, config: Dict[str, Any]) -> Dict[str, Any]:
        """Process a single worksheet according to configuration."""
        sheet_data = {}
        
        if 'subtables' not in config:
            raise ValidationError("Sheet configuration missing 'subtables'")
        
        for subtable_config in config['subtables']:
            subtable_name = subtable_config.get('name', 'unnamed_subtable')
            logger.debug(f"Processing subtable: {subtable_name}")
            
            try:
                subtable_data = self._extract_subtable(worksheet, subtable_config)
                sheet_data[subtable_name] = subtable_data
            except Exception as e:
                logger.error(f"Failed to process subtable '{subtable_name}': {e}")
                sheet_data[subtable_name] = {}
        
        return sheet_data
    
    def _extract_subtable(self, worksheet: Worksheet, config: Dict[str, Any]) -> Dict[str, Any]:
        """Extract data for a specific subtable configuration."""
        subtable_type = config.get('type', 'table')
        header_search = config.get('header_search', {})
        data_extraction = config.get('data_extraction', {})
        
        # Find the header location
        header_location = self._find_header_location(worksheet, header_search)
        if not header_location:
            logger.warning("Header location not found")
            return {}
        
        # Extract data based on type
        if subtable_type == 'key_value_pairs':
            return self._extract_key_value_pairs(worksheet, header_location, data_extraction)
        elif subtable_type == 'table':
            return self._extract_table_data(worksheet, header_location, data_extraction)
        else:
            raise ValidationError(f"Unknown subtable type: {subtable_type}")
    
    def _find_header_location(self, worksheet: Worksheet, search_config: Dict[str, Any]) -> Optional[Tuple[int, int]]:
        """Find header location based on search configuration."""
        method = search_config.get('method', 'contains_text')
        search_text = search_config.get('text', '')
        search_column = search_config.get('column', 'A')
        search_range = search_config.get('search_range', 'A1:A10')
        
        if not validate_cell_range(search_range):
            raise ValidationError(f"Invalid cell range: {search_range}")
        
        try:
            if method == 'contains_text':
                return self._find_by_text_contains(worksheet, search_text, search_range)
            elif method == 'exact_match':
                return self._find_by_exact_match(worksheet, search_text, search_range)
            elif method == 'regex':
                return self._find_by_regex(worksheet, search_text, search_range)
            else:
                raise ValidationError(f"Unknown search method: {method}")
        
        except Exception as e:
            logger.error(f"Header search failed: {e}")
            return None
    
    def _find_by_text_contains(self, worksheet: Worksheet, search_text: str, search_range: str) -> Optional[Tuple[int, int]]:
        """Find cell containing specific text."""
        for row in worksheet[search_range]:
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    if search_text.lower() in cell.value.lower():
                        return (cell.row, cell.column)
        return None
    
    def _find_by_exact_match(self, worksheet: Worksheet, search_text: str, search_range: str) -> Optional[Tuple[int, int]]:
        """Find cell with exact text match."""
        for row in worksheet[search_range]:
            for cell in row:
                if cell.value and str(cell.value).strip() == search_text.strip():
                    return (cell.row, cell.column)
        return None
    
    def _find_by_regex(self, worksheet: Worksheet, pattern: str, search_range: str) -> Optional[Tuple[int, int]]:
        """Find cell matching regex pattern."""
        try:
            regex = re.compile(pattern)
            for row in worksheet[search_range]:
                for cell in row:
                    if cell.value and regex.search(str(cell.value)):
                        return (cell.row, cell.column)
        except re.error as e:
            raise ValidationError(f"Invalid regex pattern: {e}")
        return None
    
    def _extract_key_value_pairs(self, worksheet: Worksheet, header_location: Tuple[int, int], config: Dict[str, Any]) -> Dict[str, Any]:
        """Extract key-value pairs from Excel sheet."""
        header_row, header_col = header_location
        orientation = config.get('orientation', 'horizontal')
        max_pairs = config.get('max_columns', 10) if orientation == 'horizontal' else config.get('max_rows', 10)
        column_mappings = config.get('column_mappings', {})
        
        data = {}
        
        try:
            if orientation == 'horizontal':
                # Keys in one row, values in the next row
                keys_row = header_row + config.get('headers_row_offset', 0)
                values_row = header_row + config.get('data_row_offset', 1)
                
                for col_offset in range(max_pairs):
                    col = header_col + col_offset
                    key_cell = worksheet.cell(row=keys_row, column=col)
                    value_cell = worksheet.cell(row=values_row, column=col)
                    
                    if key_cell.value and not is_empty_cell_value(key_cell.value):
                        key = str(key_cell.value).strip()
                        value = value_cell.value
                        
                        # Apply column mapping if available
                        if key in column_mappings:
                            key = column_mappings[key]
                        else:
                            key = normalize_column_name(key)
                        
                        data[key] = value
            else:
                # Vertical orientation: keys in one column, values in adjacent column
                keys_col = header_col + config.get('headers_row_offset', 0)
                values_col = header_col + config.get('data_row_offset', 1)
                
                for row_offset in range(max_pairs):
                    row = header_row + row_offset
                    key_cell = worksheet.cell(row=row, column=keys_col)
                    value_cell = worksheet.cell(row=row, column=values_col)
                    
                    if key_cell.value and not is_empty_cell_value(key_cell.value):
                        key = str(key_cell.value).strip()
                        value = value_cell.value
                        
                        # Apply column mapping if available
                        if key in column_mappings:
                            key = column_mappings[key]
                        else:
                            key = normalize_column_name(key)
                        
                        data[key] = value
        
        except Exception as e:
            raise ExcelProcessingError(f"Failed to extract key-value pairs: {e}")
        
        return data
    
    def _extract_table_data(self, worksheet: Worksheet, header_location: Tuple[int, int], config: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Extract table data from Excel sheet."""
        header_row, header_col = header_location
        headers_row = header_row + config.get('headers_row_offset', 0)
        data_start_row = headers_row + config.get('data_row_offset', 1)
        max_columns = config.get('max_columns', 20)
        max_rows = config.get('max_rows', 1000)
        column_mappings = config.get('column_mappings', {})
        
        try:
            # Extract headers
            headers = []
            for col_offset in range(max_columns):
                col = header_col + col_offset
                header_cell = worksheet.cell(row=headers_row, column=col)
                
                if header_cell.value and not is_empty_cell_value(header_cell.value):
                    header = str(header_cell.value).strip()
                    if header in column_mappings:
                        header = column_mappings[header]
                    else:
                        header = normalize_column_name(header)
                    headers.append(header)
                else:
                    break  # Stop when we hit an empty header
            
            if not headers:
                logger.warning("No headers found for table extraction")
                return []
            
            # Extract data rows
            data_rows = []
            for row_offset in range(max_rows):
                row = data_start_row + row_offset
                row_data = {}
                has_data = False
                
                for col_offset, header in enumerate(headers):
                    col = header_col + col_offset
                    cell = worksheet.cell(row=row, column=col)
                    value = cell.value
                    
                    if not is_empty_cell_value(value):
                        has_data = True
                    
                    row_data[header] = value
                
                if not has_data:
                    # Empty row found, stop extraction
                    break
                
                data_rows.append(row_data)
            
            return data_rows
        
        except Exception as e:
            raise ExcelProcessingError(f"Failed to extract table data: {e}")
    
    def extract_images(self, output_directory: str) -> Dict[str, List[str]]:
        """Extract images from Excel file and save to directory."""
        if not self.workbook:
            self._validate_file()
        
        images = {}
        
        try:
            for sheet_name in self.workbook.sheetnames:
                worksheet = self.workbook[sheet_name]
                sheet_images = []
                
                # Extract images from worksheet
                if hasattr(worksheet, '_images'):
                    for idx, img in enumerate(worksheet._images):
                        try:
                            # Get image data
                            image_data = img.ref
                            if hasattr(image_data, 'data'):
                                img_bytes = image_data.data
                            else:
                                continue
                            
                            # Convert to PIL Image
                            pil_image = PILImage.open(io.BytesIO(img_bytes))
                            
                            # Generate filename
                            filename = f"{sheet_name}_image_{idx + 1}.png"
                            filepath = os.path.join(output_directory, filename)
                            
                            # Save image
                            pil_image.save(filepath, format='PNG')
                            sheet_images.append(filepath)
                            
                            logger.debug(f"Extracted image: {filepath}")
                            
                        except Exception as e:
                            logger.warning(f"Failed to extract image {idx} from {sheet_name}: {e}")
                
                if sheet_images:
                    images[sheet_name] = sheet_images
            
            return images
        
        except Exception as e:
            raise ExcelProcessingError(f"Failed to extract images: {e}")
    
    def get_cell_value(self, sheet_name: str, cell_reference: str) -> Any:
        """Get value from specific cell."""
        if not self.workbook:
            self._validate_file()
        
        try:
            if sheet_name not in self.workbook.sheetnames:
                raise ValidationError(f"Sheet '{sheet_name}' not found")
            
            worksheet = self.workbook[sheet_name]
            cell = worksheet[cell_reference]
            return cell.value
        
        except Exception as e:
            raise ExcelProcessingError(f"Failed to get cell value: {e}")
    
    def get_range_values(self, sheet_name: str, cell_range: str) -> List[List[Any]]:
        """Get values from cell range."""
        if not self.workbook:
            self._validate_file()
        
        try:
            if sheet_name not in self.workbook.sheetnames:
                raise ValidationError(f"Sheet '{sheet_name}' not found")
            
            if not validate_cell_range(cell_range):
                raise ValidationError(f"Invalid cell range: {cell_range}")
            
            worksheet = self.workbook[sheet_name]
            cell_range_obj = worksheet[cell_range]
            
            # Handle single cell vs range
            if hasattr(cell_range_obj, '__iter__') and not isinstance(cell_range_obj, str):
                return [[cell.value for cell in row] for row in cell_range_obj]
            else:
                return [[cell_range_obj.value]]
        
        except Exception as e:
            raise ExcelProcessingError(f"Failed to get range values: {e}")
    
    def to_dataframe(self, sheet_name: str, header_row: Optional[int] = None) -> pd.DataFrame:
        """Convert Excel sheet to pandas DataFrame."""
        try:
            df = pd.read_excel(
                self.file_path,
                sheet_name=sheet_name,
                header=header_row
            )
            return df
        
        except Exception as e:
            raise ExcelProcessingError(f"Failed to convert to DataFrame: {e}")
    
    def close(self) -> None:
        """Close the workbook and free resources."""
        if self.workbook:
            self.workbook.close()
            self.workbook = None