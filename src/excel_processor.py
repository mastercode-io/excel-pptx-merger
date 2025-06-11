"""Excel data extraction and processing module with enhanced image position extraction."""

import logging
import os
import re
from typing import Any, Dict, List, Optional, Tuple, Union
import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter, column_index_from_string
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

                # Normalize sheet name for JSON compatibility
                normalized_sheet_name = normalize_column_name(sheet_name)
                logger.debug(f"Normalized sheet name: '{sheet_name}' -> '{normalized_sheet_name}'")

                extracted_data[normalized_sheet_name] = sheet_data

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

    def extract_images(self, output_directory: str) -> Dict[str, List[Dict[str, Any]]]:
        """Extract images from Excel file with enhanced position information."""
        if not self.workbook:
            self._validate_file()

        # Ensure output directory exists
        os.makedirs(output_directory, exist_ok=True)

        images = {}

        try:
            for sheet_name in self.workbook.sheetnames:
                worksheet = self.workbook[sheet_name]
                sheet_images = []

                # Extract images from worksheet
                if hasattr(worksheet, '_images') and worksheet._images:
                    for idx, img in enumerate(worksheet._images):
                        try:
                            # Get image data (corrected method)
                            image_data = img._data()

                            # Convert to PIL Image to get format and size
                            image_stream = io.BytesIO(image_data)
                            pil_image = PILImage.open(image_stream)
                            format_lower = pil_image.format.lower() if pil_image.format else 'png'

                            # Generate normalized filename
                            filename = self._normalize_image_filename(sheet_name, idx + 1, format_lower)
                            filepath = os.path.join(output_directory, filename)

                            # Extract position information
                            position_info = self._extract_image_position(img, worksheet)

                            # Save image
                            with open(filepath, 'wb') as f:
                                f.write(image_data)

                            # Create enhanced image info
                            image_info = {
                                'path': filepath,
                                'filename': filename,
                                'index': idx + 1,
                                'sheet': sheet_name,
                                'position': position_info,
                                'size': {
                                    'width': pil_image.width,
                                    'height': pil_image.height
                                },
                                'format': format_lower
                            }

                            sheet_images.append(image_info)
                            logger.debug(f"Extracted image with position: {filepath}")

                        except Exception as e:
                            logger.warning(f"Failed to extract image {idx} from {sheet_name}: {e}")
                            continue

                if sheet_images:
                    images[sheet_name] = sheet_images

            return images

        except Exception as e:
            raise ExcelProcessingError(f"Failed to extract images: {e}")

    def _normalize_image_filename(self, sheet_name: str, idx: int, format_name: str) -> str:
        """Normalize image filename for consistency and reliability."""
        # Normalize sheet name: lowercase, replace spaces with underscores, remove special chars
        normalized_sheet = re.sub(r'[^\w_]', '', sheet_name.lower().replace(' ', '_'))
        # Ensure format is lowercase and without leading dot
        format_lower = format_name.lower().lstrip('.')
        # Create normalized filename
        return f"{normalized_sheet}_image_{idx}.{format_lower}"

    def _extract_image_position(self, img, worksheet: Worksheet) -> Dict[str, Any]:
        """Extract position information from image object."""
        position_info = {
            'anchor_type': None,
            'from_cell': None,
            'to_cell': None,
            'coordinates': None,
            'estimated_cell': None,
            'size_info': None
        }

        try:
            if hasattr(img, 'anchor'):
                anchor = img.anchor
                position_info['anchor_type'] = type(anchor).__name__

                # Two-cell anchor (most common)
                if hasattr(anchor, '_from') and hasattr(anchor, 'to'):
                    from_info = anchor._from
                    to_info = anchor.to

                    # Convert to Excel cell references
                    from_col = get_column_letter(from_info.col + 1)
                    from_cell = f"{from_col}{from_info.row + 1}"
                    to_col = get_column_letter(to_info.col + 1)
                    to_cell = f"{to_col}{to_info.row + 1}"

                    position_info.update({
                        'from_cell': from_cell,
                        'to_cell': to_cell,
                        'coordinates': {
                            'from': {'col': from_info.col, 'row': from_info.row},
                            'to': {'col': to_info.col, 'row': to_info.row}
                        }
                    })

                    # Use from_cell as the primary estimated position
                    position_info['estimated_cell'] = from_cell

                    # Calculate span information
                    col_span = to_info.col - from_info.col + 1
                    row_span = to_info.row - from_info.row + 1
                    position_info['size_info'] = {
                        'column_span': col_span,
                        'row_span': row_span
                    }

                # One-cell anchor
                elif hasattr(anchor, '_from'):
                    from_info = anchor._from
                    from_col = get_column_letter(from_info.col + 1)
                    from_cell = f"{from_col}{from_info.row + 1}"

                    position_info.update({
                        'from_cell': from_cell,
                        'estimated_cell': from_cell,
                        'coordinates': {
                            'from': {'col': from_info.col, 'row': from_info.row}
                        }
                    })

                    position_info['size_info'] = {
                        'column_span': 1,
                        'row_span': 1
                    }

        except Exception as e:
            logger.debug(f"Could not extract detailed position info: {e}")
            # Fallback to basic position estimation
            position_info['estimated_cell'] = 'A1'  # Default fallback
            position_info['size_info'] = {'column_span': 1, 'row_span': 1}

        return position_info

    def get_image_by_position(self, images: Dict[str, List[Dict[str, Any]]],
                             target_cell: str, sheet_name: Optional[str] = None) -> Optional[Dict[str, Any]]:
        """Find image by cell position or proximity."""
        sheets_to_search = [sheet_name] if sheet_name else images.keys()

        for sheet in sheets_to_search:
            if sheet not in images:
                continue

            for image_info in images[sheet]:
                position = image_info.get('position', {})

                # Exact match on from_cell or estimated_cell
                if (position.get('from_cell') == target_cell or
                    position.get('estimated_cell') == target_cell):
                    return image_info

                # Check if target position falls within image range
                if position.get('coordinates'):
                    coords = position['coordinates']
                    if 'from' in coords and 'to' in coords:
                        target_coords = self._cell_to_coordinates(target_cell)
                        if target_coords and self._is_in_range(target_coords, coords):
                            return image_info

        return None

    def _cell_to_coordinates(self, cell_ref: str) -> Optional[Dict[str, int]]:
        """Convert cell reference like 'A1' to coordinates."""
        try:
            match = re.match(r'^([A-Z]+)(\d+)$', cell_ref.upper())
            if match:
                col_str, row_str = match.groups()
                return {
                    'col': column_index_from_string(col_str) - 1,  # 0-based
                    'row': int(row_str) - 1  # 0-based
                }
        except Exception as e:
            logger.debug(f"Failed to convert cell reference {cell_ref}: {e}")
        return None

    def _is_in_range(self, target: Dict[str, int], range_coords: Dict) -> bool:
        """Check if target coordinates fall within the given range."""
        try:
            from_coords = range_coords['from']
            to_coords = range_coords['to']

            return (from_coords['col'] <= target['col'] <= to_coords['col'] and
                    from_coords['row'] <= target['row'] <= to_coords['row'])
        except (KeyError, TypeError):
            return False

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

    def get_image_summary(self, images: Dict[str, List[Dict[str, Any]]]) -> Dict[str, Any]:
        """Get summary information about extracted images."""
        summary = {
            'total_images': 0,
            'sheets_with_images': 0,
            'images_by_sheet': {},
            'position_summary': {
                'with_position': 0,
                'without_position': 0,
                'position_types': {}
            }
        }

        for sheet_name, sheet_images in images.items():
            if sheet_images:
                summary['sheets_with_images'] += 1
                summary['images_by_sheet'][sheet_name] = len(sheet_images)
                summary['total_images'] += len(sheet_images)

                for image_info in sheet_images:
                    position = image_info.get('position', {})
                    if position.get('estimated_cell'):
                        summary['position_summary']['with_position'] += 1
                        anchor_type = position.get('anchor_type', 'unknown')
                        summary['position_summary']['position_types'][anchor_type] = \
                            summary['position_summary']['position_types'].get(anchor_type, 0) + 1
                    else:
                        summary['position_summary']['without_position'] += 1

        return summary
    
    def close(self) -> None:
        """Close the workbook and free resources."""
        if self.workbook:
            self.workbook.close()
            self.workbook = None
