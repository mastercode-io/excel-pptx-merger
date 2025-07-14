"""Excel file update module for modifying cells with new data."""

import logging
import os
import io
import uuid
import re
import base64
import requests
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple, Union
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.cell import coordinate_from_string
from PIL import Image as PILImage

from .temp_file_manager import TempFileManager
from .utils.exceptions import ExcelProcessingError, ValidationError

logger = logging.getLogger(__name__)


class ExcelUpdater:
    """Updates Excel files with new data according to configuration."""

    def __init__(self, excel_file_path: str):
        """Initialize Excel updater with file path."""
        self.file_path = excel_file_path
        # Load workbook with minimal parameters to avoid corruption
        # Use default settings that are most compatible
        self.workbook = load_workbook(excel_file_path)
        self.update_log = []
        self.temp_image_files = []  # Track temp files for cleanup

    def update_excel(
        self,
        update_data: Dict[str, Any],
        config: Dict[str, Any],
        include_update_log: bool = False,
    ) -> str:
        """
        Main update method - returns path to updated file.

        Args:
            update_data: Dictionary with data to update
            config: Configuration mapping for updates
            include_update_log: Whether to include diagnostic update_log sheet (default: False)

        Returns:
            Path to updated Excel file
        """
        try:
            self._log_info("Starting Excel update process")

            # Check existing content for preservation
            self._log_existing_content()

            # Validate configuration
            self._validate_update_config(config)

            # Get list of sheets that will be updated
            sheets_to_update = set(config.get("sheet_configs", {}).keys())
            self._log_info(f"Sheets to be updated: {sheets_to_update}")

            # Process each sheet
            for sheet_name, sheet_config in config.get("sheet_configs", {}).items():
                if sheet_name not in self.workbook.sheetnames:
                    self._log_error(f"Sheet '{sheet_name}' not found in workbook")
                    continue

                sheet = self.workbook[sheet_name]
                self._log_info(f"Processing sheet: {sheet_name}")

                # Process each subtable in the sheet
                for subtable_config in sheet_config.get("subtables", []):
                    self._update_subtable(sheet, subtable_config, update_data)

            # Verify non-updated sheets remain intact
            for sheet_name in self.workbook.sheetnames:
                if sheet_name not in sheets_to_update and sheet_name != "update_log":
                    self._verify_sheet_preservation(sheet_name)

            # Add diagnostic log sheet if requested
            if include_update_log:
                self._add_update_log_sheet()

            # Save updated workbook
            output_path = self._save_updated_workbook()
            self._log_info(f"Excel update completed successfully: {output_path}")

            return output_path

        except Exception as e:
            self._log_error(f"Update process failed: {e}")
            raise ExcelProcessingError(f"Failed to update Excel file: {e}")

    def _update_subtable(
        self,
        sheet: Worksheet,
        subtable_config: Dict[str, Any],
        update_data: Dict[str, Any],
    ) -> None:
        """Update subtable with unified offset handling."""
        subtable_name = subtable_config["name"]
        subtable_type = subtable_config["type"]

        self._log_info(f"Updating subtable: {subtable_name}")

        # Skip if no data_update configuration
        if "data_update" not in subtable_config:
            self._log_warning(
                f"Skipping subtable '{subtable_name}' - no data_update configuration"
            )
            return

        # Find starting location
        location = self._find_update_location(sheet, subtable_config["header_search"])
        if not location["found"]:
            self._log_error(f"Could not find location for subtable '{subtable_name}'")
            return

        # Get data for this subtable
        if subtable_name not in update_data:
            self._log_warning(f"No data provided for subtable '{subtable_name}'")
            return

        data = update_data[subtable_name]
        update_config = subtable_config["data_update"]

        # Apply offsets to starting location
        base_row = location["row"] + update_config.get("headers_row_offset", 0)
        base_col = location["col"] + update_config.get("headers_col_offset", 0)

        # Update based on subtable type
        if subtable_type == "key_value_pairs":
            self._update_key_value_pairs_with_offsets(
                sheet, base_row, base_col, update_config, data
            )
        elif subtable_type == "table":
            self._update_table_with_offsets(
                sheet, base_row, base_col, update_config, data
            )
        elif subtable_type == "matrix_table":
            self._update_matrix_table_with_offsets(
                sheet, base_row, base_col, update_config, data
            )

    def _find_update_location(
        self, sheet: Worksheet, header_search_config: Dict[str, Any]
    ) -> Dict[str, Any]:
        """Find starting location using either method."""
        method = header_search_config["method"]

        if method == "cell_address":
            return self._find_by_cell_address(sheet, header_search_config)
        elif method == "contains_text":
            return self._find_by_contains_text(sheet, header_search_config)
        else:
            raise ValueError(f"Unsupported header search method: {method}")

    def _find_by_cell_address(
        self, sheet: Worksheet, config: Dict[str, Any]
    ) -> Dict[str, Any]:
        """Direct cell address method."""
        cell_address = config["cell"]
        try:
            row, col = self._parse_cell_address(cell_address)
            self._log_info(
                f"Found location by cell address {cell_address}: row={row}, col={col}"
            )
            return {"row": row, "col": col, "found": True, "address": cell_address}
        except Exception as e:
            self._log_error(f"Invalid cell address {cell_address}: {e}")
            return {"found": False}

    def _find_by_contains_text(
        self, sheet: Worksheet, config: Dict[str, Any]
    ) -> Dict[str, Any]:
        """Text search method."""
        search_text = config["text"]
        search_column = config["column"]
        search_range = config.get(
            "search_range", f"{search_column}1:{search_column}100"
        )

        try:
            for cell in sheet[search_range]:
                for cell_obj in cell:
                    if (
                        cell_obj.value
                        and search_text.lower() in str(cell_obj.value).lower()
                    ):
                        self._log_info(
                            f"Found location by text '{search_text}' at {cell_obj.coordinate}"
                        )
                        return {
                            "row": cell_obj.row,
                            "col": cell_obj.column,
                            "found": True,
                            "address": cell_obj.coordinate,
                        }

            self._log_error(f"Text '{search_text}' not found in range {search_range}")
            return {"found": False}

        except Exception as e:
            self._log_error(f"Text search failed: {e}")
            return {"found": False}

    def _update_key_value_pairs_with_offsets(
        self,
        sheet: Worksheet,
        base_row: int,
        base_col: int,
        config: Dict[str, Any],
        data: Dict[str, Any],
    ) -> None:
        """Update key-value pairs with offset support."""
        column_mappings = config["column_mappings"]
        orientation = config.get("orientation", "horizontal")
        data_row_offset = config.get("data_row_offset", 1)
        data_col_offset = config.get("data_col_offset", 1)

        if orientation == "horizontal":
            # Headers are in base_row, data in base_row + data_row_offset
            data_row = base_row + data_row_offset

            for mapping_key, field_config in column_mappings.items():
                field_name = field_config["name"]
                field_type = field_config["type"]

                if field_name in data:
                    # Handle both cell_address and column-based mappings
                    if self._is_cell_address(mapping_key):
                        # Direct cell address (e.g., "B14")
                        cell_address = mapping_key
                    else:
                        # Column header name - find the column
                        header_col = self._find_header_column(
                            sheet, base_row, mapping_key, config, base_col
                        )
                        if header_col:
                            cell_address = f"{get_column_letter(header_col)}{data_row}"
                        else:
                            self._log_error(f"Header '{mapping_key}' not found")
                            continue

                    self._update_cell(sheet, cell_address, data[field_name], field_type)

        elif orientation == "vertical":
            # Keys are in base_col, data in base_col + data_col_offset
            data_col = base_col + data_col_offset

            for mapping_key, field_config in column_mappings.items():
                field_name = field_config["name"]
                field_type = field_config["type"]

                if field_name in data:
                    # Find the row where this key is located
                    key_row = self._find_key_row(
                        sheet, base_row, base_col, mapping_key, config
                    )
                    if key_row:
                        cell_address = f"{get_column_letter(data_col)}{key_row}"
                        self._update_cell(
                            sheet, cell_address, data[field_name], field_type
                        )
                    else:
                        self._log_error(
                            f"Key '{mapping_key}' not found in vertical layout"
                        )

    def _update_table_with_offsets(
        self,
        sheet: Worksheet,
        base_row: int,
        base_col: int,
        config: Dict[str, Any],
        data: List[Dict[str, Any]],
    ) -> None:
        """Update table data with offset support."""
        column_mappings = config["column_mappings"]
        orientation = config.get("orientation", "vertical")
        data_row_offset = config.get("data_row_offset", 1)

        if orientation == "vertical":
            # Start data from base_row + data_row_offset
            data_start_row = base_row + data_row_offset

            for row_idx, row_data in enumerate(data):
                current_row = data_start_row + row_idx

                for mapping_key, field_config in column_mappings.items():
                    field_name = field_config["name"]
                    field_type = field_config["type"]

                    if field_name in row_data:
                        # Handle both column letters and header names
                        if self._is_column_letter(mapping_key):
                            # Direct column (e.g., "A")
                            cell_address = f"{mapping_key}{current_row}"
                        else:
                            # Header name - find the column
                            header_col = self._find_header_column(
                                sheet, base_row, mapping_key, config, base_col
                            )
                            if header_col:
                                cell_address = (
                                    f"{get_column_letter(header_col)}{current_row}"
                                )
                            else:
                                self._log_error(f"Header '{mapping_key}' not found")
                                continue

                        self._update_cell(
                            sheet, cell_address, row_data[field_name], field_type
                        )

    def _update_matrix_table_with_offsets(
        self,
        sheet: Worksheet,
        base_row: int,
        base_col: int,
        config: Dict[str, Any],
        data: Dict[str, Dict[str, Any]],
    ) -> None:
        """Update matrix table data with offset support.

        Expects data in format: {row_key: {col_key: value}}
        """
        column_mappings = config["column_mappings"]
        row_key_mappings = config.get("row_key_mappings", {})

        headers_row = base_row + config.get("headers_row_offset", 0)
        data_start_row = headers_row + config.get("data_row_offset", 1)

        # Support for column offset - allows table to start in different column than search text
        headers_col_offset = config.get("headers_col_offset", 0)
        header_col = base_col + headers_col_offset

        # Row keys are in the first column of the data area
        row_keys_col_offset = config.get("row_keys_col_offset", 0)
        row_keys_col = header_col + row_keys_col_offset

        # Data starts in the column after row keys
        data_col_offset = config.get("data_col_offset", 1)
        data_start_col = row_keys_col + data_col_offset

        max_rows = config.get("max_rows", 1000)

        # Create reverse mapping for row keys (JSON key -> Excel key)
        reverse_row_key_mappings = {}
        for excel_key, json_key in row_key_mappings.items():
            reverse_row_key_mappings[json_key] = excel_key

        # Create reverse mapping for column headers (JSON key -> Excel key)
        reverse_column_mappings = {}
        for excel_key, mapping in column_mappings.items():
            if isinstance(mapping, str):
                json_key = mapping
            else:
                json_key = mapping.get("name")
            if json_key:
                reverse_column_mappings[json_key] = excel_key

        # Update data for each row
        for json_row_key, row_data in data.items():
            if json_row_key == "_field_types":  # Skip metadata
                continue

            # Find the Excel row for this row key
            target_row = None

            # Look for the row key in the Excel sheet
            for row_offset in range(max_rows):
                row = data_start_row + row_offset
                row_key_cell = sheet.cell(row=row, column=row_keys_col)

                if row_key_cell.value:
                    excel_row_key = str(row_key_cell.value).strip()

                    # Check if this matches our target row key (either directly or through mapping)
                    if (
                        excel_row_key == json_row_key
                        or reverse_row_key_mappings.get(json_row_key) == excel_row_key
                    ):
                        target_row = row
                        break

            if target_row is None:
                self._log_error(f"Row key '{json_row_key}' not found in matrix table")
                continue

            # Update cells in this row
            for json_col_key, value in row_data.items():
                # Find the Excel column for this column key
                target_col = None

                # Get the column header mapping
                excel_col_key = reverse_column_mappings.get(json_col_key, json_col_key)

                # Find the column in the header row
                header_col_found = self._find_header_column(
                    sheet, headers_row, excel_col_key, config, data_start_col
                )

                if header_col_found:
                    target_col = header_col_found
                else:
                    self._log_error(
                        f"Column '{json_col_key}' (Excel: '{excel_col_key}') not found in matrix table"
                    )
                    continue

                # Get field type from column mappings
                if excel_col_key in column_mappings:
                    mapping = column_mappings[excel_col_key]
                    if isinstance(mapping, str):
                        field_type = "text"
                    else:
                        field_type = mapping.get("type", "text")
                else:
                    field_type = "text"

                # Update the cell
                cell_address = f"{get_column_letter(target_col)}{target_row}"
                self._update_cell(sheet, cell_address, value, field_type)

    def _update_cell(
        self, sheet: Worksheet, cell_address: str, value: Any, field_type: str
    ) -> bool:
        """Update single cell with type-specific handling."""
        try:
            cell = sheet[cell_address]
            original_value = cell.value

            if field_type == "text":
                cell.value = str(value) if value is not None else ""
                self._log_success(f"Updated {cell_address} with text: '{value}'")

            elif field_type == "number":
                try:
                    cell.value = float(value) if value is not None else 0
                    self._log_success(f"Updated {cell_address} with number: {value}")
                except (ValueError, TypeError):
                    cell.value = str(value) if value is not None else ""
                    self._log_warning(
                        f"Could not convert '{value}' to number, saved as text in {cell_address}"
                    )

            elif field_type == "image":
                return self._insert_image_v2(sheet, cell_address, value)

            else:
                # Default to text
                cell.value = str(value) if value is not None else ""
                self._log_success(
                    f"Updated {cell_address} with default text: '{value}'"
                )

            return True

        except Exception as e:
            self._log_error(f"Failed to update {cell_address}: {e}")
            try:
                # Try to set error value if valid cell address
                if self._is_cell_address(cell_address):
                    sheet[cell_address].value = "!#VALUE"
            except:
                pass  # If even error setting fails, just continue
            return False

    def _insert_image(
        self, sheet: Worksheet, cell_address: str, image_data: Any
    ) -> bool:
        """Insert image above specified cell."""
        temp_image_path = None
        try:
            if not image_data:
                sheet[cell_address].value = "[NO_IMAGE_DATA]"
                self._log_warning(f"No image data provided for {cell_address}")
                return True

            # Handle different image input types
            final_image_path = None
            final_size = None

            if isinstance(image_data, str):
                if image_data.startswith("file://") or os.path.exists(image_data):
                    # Local file path - try to use directly without processing
                    try:
                        file_path = image_data
                        if file_path.startswith("file://"):
                            import urllib.parse

                            file_path = urllib.parse.unquote(file_path[7:])

                        if os.path.exists(file_path):
                            # Load and process for sizing only, but use original file for Excel
                            with open(file_path, "rb") as f:
                                image_bytes = f.read()
                            pil_img = PILImage.open(io.BytesIO(image_bytes))
                            processed_img, final_size = self._process_image_with_sizing(
                                pil_img
                            )

                            # If image needs resizing, create temp file
                            if final_size != pil_img.size:
                                import tempfile

                                temp_fd, temp_image_path = tempfile.mkstemp(
                                    suffix=".png"
                                )
                                os.close(temp_fd)
                                processed_img.save(temp_image_path, format="PNG")
                                self.temp_image_files.append(temp_image_path)
                                final_image_path = temp_image_path
                            else:
                                # Use original file directly
                                final_image_path = file_path
                        else:
                            raise ValueError(f"File not found: {file_path}")
                    except Exception as e:
                        self._log_error(
                            f"Failed to process local file {image_data}: {e}"
                        )
                        sheet[cell_address].value = "!#VALUE"
                        return False

                elif image_data.startswith("data:image"):
                    # Base64 image - decode and process
                    image_bytes = self._decode_base64_image(image_data)
                    pil_img = PILImage.open(io.BytesIO(image_bytes))
                    processed_img, final_size = self._process_image_with_sizing(pil_img)

                    import tempfile

                    temp_fd, temp_image_path = tempfile.mkstemp(suffix=".png")
                    os.close(temp_fd)
                    processed_img.save(temp_image_path, format="PNG")
                    self.temp_image_files.append(temp_image_path)
                    final_image_path = temp_image_path

                elif image_data.startswith("http"):
                    # Download URL and process
                    image_bytes = self._download_image(image_data)
                    pil_img = PILImage.open(io.BytesIO(image_bytes))
                    processed_img, final_size = self._process_image_with_sizing(pil_img)

                    import tempfile

                    temp_fd, temp_image_path = tempfile.mkstemp(suffix=".png")
                    os.close(temp_fd)
                    processed_img.save(temp_image_path, format="PNG")
                    self.temp_image_files.append(temp_image_path)
                    final_image_path = temp_image_path
                else:
                    # Text fallback
                    sheet[cell_address].value = str(image_data)
                    self._log_info(
                        f"Inserted text instead of image at {cell_address}: {image_data}"
                    )
                    return True
            else:
                # Unknown format - use text
                sheet[cell_address].value = str(image_data)
                self._log_warning(
                    f"Unknown image format, inserted as text at {cell_address}"
                )
                return True

            if not final_image_path or not final_size:
                sheet[cell_address].value = "!#VALUE"
                self._log_error(f"Could not process image data for {cell_address}")
                return False

            # Create Excel image from final path
            excel_img = ExcelImage(final_image_path)

            # Position at target cell
            row, col = self._parse_cell_address(cell_address)
            excel_img.anchor = f"{get_column_letter(col)}{row}"

            sheet.add_image(excel_img)
            # Keep cell empty - no text pollution

            # Adjust row height to fit image
            self._adjust_row_height(sheet, row, final_size[1])

            self._log_success(
                f"Inserted image at {cell_address} (size: {final_size[0]}x{final_size[1]})"
            )
            return True

        except Exception as e:
            self._log_error(f"Image insert failed at {cell_address}: {e}")
            try:
                # Try to set error value if valid cell address
                if self._is_cell_address(cell_address):
                    sheet[cell_address].value = "!#VALUE"
            except:
                pass  # If even error setting fails, just continue
            return False

    def _decode_base64_image(self, base64_data: str) -> bytes:
        """Decode base64 image data."""
        try:
            # Extract base64 data after comma (remove data:image/png;base64, prefix)
            if "," in base64_data:
                base64_data = base64_data.split(",", 1)[1]

            return base64.b64decode(base64_data)
        except Exception as e:
            raise ValueError(f"Invalid base64 image data: {e}")

    def _download_image(self, url: str, timeout: int = 10) -> bytes:
        """Download image from URL with timeout."""
        try:
            response = requests.get(url, timeout=timeout, stream=True)
            response.raise_for_status()

            # Check content type
            content_type = response.headers.get("content-type", "")
            if not content_type.startswith("image/"):
                raise ValueError(
                    f"URL does not point to an image (content-type: {content_type})"
                )

            return response.content

        except Exception as e:
            raise ValueError(f"Failed to download image from {url}: {e}")

    def _load_local_image(self, file_path: str) -> bytes:
        """Load image from local file path."""
        try:
            # Handle file:// URLs
            if file_path.startswith("file://"):
                import urllib.parse

                file_path = urllib.parse.unquote(
                    file_path[7:]
                )  # Remove 'file://' prefix

            # Check if file exists
            if not os.path.exists(file_path):
                raise ValueError(f"Image file not found: {file_path}")

            # Check if it's a valid image file
            valid_extensions = {
                ".png",
                ".jpg",
                ".jpeg",
                ".gif",
                ".bmp",
                ".tiff",
                ".webp",
            }
            file_ext = os.path.splitext(file_path.lower())[1]
            if file_ext not in valid_extensions:
                raise ValueError(f"Unsupported image format: {file_ext}")

            # Read file
            with open(file_path, "rb") as f:
                return f.read()

        except Exception as e:
            raise ValueError(f"Failed to load local image from {file_path}: {e}")

    def _process_image_with_sizing(self, pil_img):
        """Process image with size management and return (processed_image, (width, height))."""
        try:
            # Default image size in points (Excel units)
            default_width = 50  # points
            default_height = 25  # points

            # Get original size
            original_width, original_height = pil_img.size

            # Calculate if we need to resize (if image is larger than default)
            # Convert points to pixels for comparison (1 point ≈ 1.33 pixels)
            max_width_pixels = int(default_width * 1.33)
            max_height_pixels = int(default_height * 1.33)

            if (
                original_width <= max_width_pixels
                and original_height <= max_height_pixels
            ):
                # Image is small enough, use original size (copy to new image to avoid fp issues)
                self._log_info(
                    f"Using original image size: {original_width}x{original_height}px"
                )
                # Create a new image to avoid file pointer issues
                new_img = pil_img.copy()
                return new_img, (original_width, original_height)
            else:
                # Resize to default size while maintaining aspect ratio
                new_img = pil_img.copy()
                new_img.thumbnail(
                    (max_width_pixels, max_height_pixels), PILImage.Resampling.LANCZOS
                )
                new_size = new_img.size
                self._log_info(
                    f"Resized image from {original_width}x{original_height}px to {new_size[0]}x{new_size[1]}px"
                )
                return new_img, new_size

        except Exception as e:
            self._log_error(f"Image processing failed: {e}")
            # Return copy of original image if processing fails
            return pil_img.copy(), pil_img.size

    def _get_image_dimensions_lightweight(self, image_path_or_bytes):
        """Get image dimensions without full PIL processing to avoid corruption."""
        try:
            if isinstance(image_path_or_bytes, str):
                # File path
                with open(image_path_or_bytes, "rb") as f:
                    image_bytes = f.read()
            else:
                # Already bytes
                image_bytes = image_path_or_bytes

            # Use PIL but only for dimensions, then immediately close
            temp_img = PILImage.open(io.BytesIO(image_bytes))
            dimensions = temp_img.size
            temp_img.close()  # Immediately close to avoid corruption

            return dimensions

        except Exception as e:
            self._log_warning(f"Failed to get image dimensions: {e}")
            # Return default dimensions if detection fails
            return (100, 75)  # Default fallback

    def _adjust_row_height(self, sheet, row, image_height_pixels):
        """Adjust row height to accommodate image."""
        try:
            # Convert pixels to Excel points (1 point ≈ 1.33 pixels)
            required_height_points = max(15, image_height_pixels / 1.33)

            # Get current row height (default is ~15 points)
            current_height = sheet.row_dimensions[row].height or 15

            # Only increase height if necessary
            if required_height_points > current_height:
                sheet.row_dimensions[row].height = required_height_points
                self._log_info(
                    f"Adjusted row {row} height from {current_height:.1f} to {required_height_points:.1f} points"
                )
            else:
                self._log_info(
                    f"Row {row} height sufficient: {current_height:.1f} points (needed: {required_height_points:.1f})"
                )

        except Exception as e:
            self._log_warning(f"Failed to adjust row height for row {row}: {e}")

    def _adjust_column_width(self, sheet, col, image_width_pixels):
        """Adjust column width to accommodate image."""
        try:
            # Convert pixels to Excel character widths (rough approximation)
            # Excel default column width is ~8.43 characters ≈ 64 points ≈ 85 pixels
            required_width_chars = max(
                8.43, image_width_pixels / 10
            )  # ~10 pixels per character

            # Get current column width (default is ~8.43 characters)
            col_letter = get_column_letter(col)
            current_width = sheet.column_dimensions[col_letter].width or 8.43

            # Only increase width if necessary
            if required_width_chars > current_width:
                sheet.column_dimensions[col_letter].width = required_width_chars
                self._log_info(
                    f"Adjusted column {col_letter} width from {current_width:.1f} to {required_width_chars:.1f} characters"
                )
            else:
                self._log_info(
                    f"Column {col_letter} width sufficient: {current_width:.1f} characters (needed: {required_width_chars:.1f})"
                )

        except Exception as e:
            self._log_warning(f"Failed to adjust column width for column {col}: {e}")

    def _is_cell_address(self, mapping_key: str) -> bool:
        """Check if mapping key is a cell address (e.g., 'B14')."""
        return bool(re.match(r"^[A-Z]+\d+$", mapping_key.upper()))

    def _is_column_letter(self, mapping_key: str) -> bool:
        """Check if mapping key is a column letter (e.g., 'A', 'BC')."""
        # Excel columns are max 3 letters (up to XFD = 16384 columns)
        # Common valid patterns: A, B, Z, AA, AB, ZZ, AAA, XFD
        if not mapping_key or len(mapping_key) > 3:
            return False

        # Must be all uppercase letters
        if not re.match(r"^[A-Z]+$", mapping_key.upper()):
            return False

        # Try to convert to column index - if it fails, not a valid column
        try:
            from openpyxl.utils import column_index_from_string

            column_index_from_string(mapping_key.upper())
            return True
        except ValueError:
            return False

    def _find_header_column(
        self,
        sheet: Worksheet,
        header_row: int,
        header_text: str,
        config: Dict[str, Any],
        base_col: int = 1,
    ) -> Optional[int]:
        """Find column index for header text."""
        max_columns = config.get("max_columns", 20)
        start_col = base_col

        for col_offset in range(max_columns):
            col_idx = start_col + col_offset
            cell = sheet.cell(row=header_row, column=col_idx)

            if cell.value and header_text.lower() in str(cell.value).lower():
                return col_idx

        return None

    def _find_key_row(
        self,
        sheet: Worksheet,
        base_row: int,
        base_col: int,
        key_text: str,
        config: Dict[str, Any],
    ) -> Optional[int]:
        """Find row index for key text in vertical layout."""
        max_rows = config.get("max_rows", 20)
        start_row = base_row

        for row_offset in range(max_rows):
            row_idx = start_row + row_offset
            cell = sheet.cell(row=row_idx, column=base_col)

            if cell.value and key_text.lower() in str(cell.value).lower():
                return row_idx

        return None

    def _parse_cell_address(self, cell_address: str) -> Tuple[int, int]:
        """Parse cell address into row and column."""
        try:
            col_letter, row = coordinate_from_string(cell_address)
            col_idx = column_index_from_string(col_letter)
            return row, col_idx
        except Exception:
            raise ValueError(f"Invalid cell address: {cell_address}")

    def _validate_update_config(self, config: Dict[str, Any]) -> None:
        """Validate update configuration structure."""
        if "sheet_configs" not in config:
            raise ValidationError("Configuration must contain 'sheet_configs'")

        for sheet_name, sheet_config in config.get("sheet_configs", {}).items():
            if "subtables" not in sheet_config:
                raise ValidationError(f"Sheet '{sheet_name}' must contain 'subtables'")

            for subtable in sheet_config.get("subtables", []):
                # Validate required fields (data_update OR data_extraction)
                required_base_fields = ["name", "type", "header_search"]
                for field in required_base_fields:
                    if field not in subtable:
                        raise ValidationError(
                            f"Subtable missing required field: {field}"
                        )

                # Must have either data_update or data_extraction
                if "data_update" not in subtable and "data_extraction" not in subtable:
                    raise ValidationError(
                        f"Subtable '{subtable['name']}' must have either 'data_update' or 'data_extraction'"
                    )

                # If using data_extraction, warn and skip
                if "data_extraction" in subtable and "data_update" not in subtable:
                    logger.warning(
                        f"Subtable '{subtable['name']}' has data_extraction but not data_update - skipping"
                    )

                # Validate header search method
                header_search = subtable.get("header_search", {})
                method = header_search.get("method")

                if method == "cell_address":
                    if "cell" not in header_search:
                        raise ValidationError(
                            f"cell_address method requires 'cell' field in {subtable['name']}"
                        )
                elif method == "contains_text":
                    required = ["text", "column", "search_range"]
                    for field in required:
                        if field not in header_search:
                            raise ValidationError(
                                f"contains_text method requires '{field}' field in {subtable['name']}"
                            )
                else:
                    raise ValidationError(
                        f"Unsupported method '{method}' in {subtable['name']}"
                    )

                # Validate data_update structure if present
                if "data_update" in subtable:
                    data_update = subtable.get("data_update", {})
                    if "column_mappings" not in data_update:
                        raise ValidationError(
                            f"column_mappings required in data_update for {subtable['name']}"
                        )

    def _add_update_log_sheet(self) -> None:
        """Create diagnostic sheet with update results."""
        try:
            # Remove existing log sheet if present
            if "update_log" in self.workbook.sheetnames:
                del self.workbook["update_log"]

            # Create new log sheet
            log_sheet = self.workbook.create_sheet("update_log")

            # Add headers
            headers = [
                "Timestamp",
                "Operation",
                "Cell/Range",
                "Status",
                "Details",
                "Original Value",
                "New Value",
            ]
            for col_idx, header in enumerate(headers, 1):
                log_sheet.cell(row=1, column=col_idx, value=header)

            # Add log entries
            for row_idx, log_entry in enumerate(self.update_log, 2):
                log_sheet.cell(row=row_idx, column=1, value=log_entry["timestamp"])
                log_sheet.cell(row=row_idx, column=2, value=log_entry["operation"])
                log_sheet.cell(row=row_idx, column=3, value=log_entry.get("cell", ""))
                log_sheet.cell(row=row_idx, column=4, value=log_entry["status"])
                log_sheet.cell(row=row_idx, column=5, value=log_entry["details"])
                log_sheet.cell(
                    row=row_idx, column=6, value=log_entry.get("original_value", "")
                )
                log_sheet.cell(
                    row=row_idx, column=7, value=log_entry.get("new_value", "")
                )

            self._log_info(
                f"Created update_log sheet with {len(self.update_log)} entries"
            )

        except Exception as e:
            logger.warning(f"Failed to create update log sheet: {e}")

    def _save_updated_workbook(self) -> str:
        """Save updated workbook to new file."""
        try:
            # Create output path
            base_name = os.path.splitext(os.path.basename(self.file_path))[0]
            output_dir = os.path.dirname(self.file_path)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(
                output_dir, f"updated_{base_name}_{timestamp}.xlsx"
            )

            # Save workbook
            self.workbook.save(output_path)
            return output_path
        finally:
            # Clean up temporary image files after workbook save
            self._cleanup_temp_files()

    def _log_info(self, message: str) -> None:
        """Log info message."""
        entry = {
            "timestamp": datetime.now().isoformat(),
            "operation": "INFO",
            "status": "INFO",
            "details": message,
        }
        self.update_log.append(entry)
        logger.info(message)

    def _log_success(
        self,
        message: str,
        cell: str = "",
        original_value: Any = "",
        new_value: Any = "",
    ) -> None:
        """Log success message."""
        entry = {
            "timestamp": datetime.now().isoformat(),
            "operation": "UPDATE",
            "cell": cell,
            "status": "SUCCESS",
            "details": message,
            "original_value": str(original_value) if original_value is not None else "",
            "new_value": str(new_value) if new_value is not None else "",
        }
        self.update_log.append(entry)
        logger.info(message)

    def _log_warning(self, message: str, cell: str = "") -> None:
        """Log warning message."""
        entry = {
            "timestamp": datetime.now().isoformat(),
            "operation": "WARNING",
            "cell": cell,
            "status": "WARNING",
            "details": message,
        }
        self.update_log.append(entry)
        logger.warning(message)

    def _log_error(self, message: str, cell: str = "") -> None:
        """Log error message."""
        entry = {
            "timestamp": datetime.now().isoformat(),
            "operation": "ERROR",
            "cell": cell,
            "status": "ERROR",
            "details": message,
        }
        self.update_log.append(entry)
        logger.error(message)

    def _log_existing_content(self) -> None:
        """Log all existing content for preservation verification."""
        try:
            total_images = 0
            total_charts = 0
            total_data_cells = 0

            for sheet_name in self.workbook.sheetnames:
                sheet = self.workbook[sheet_name]

                # Count images
                sheet_images = 0
                if hasattr(sheet, "_images") and sheet._images:
                    sheet_images = len(sheet._images)
                    total_images += sheet_images

                # Count charts
                sheet_charts = 0
                if hasattr(sheet, "_charts") and sheet._charts:
                    sheet_charts = len(sheet._charts)
                    total_charts += sheet_charts

                # Count data cells (non-empty cells)
                sheet_data_cells = 0
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value is not None:
                            sheet_data_cells += 1
                total_data_cells += sheet_data_cells

                self._log_info(
                    f"Sheet '{sheet_name}': {sheet_images} images, {sheet_charts} charts, {sheet_data_cells} data cells"
                )

            self._log_info(
                f"Total content: {total_images} images, {total_charts} charts, {total_data_cells} data cells"
            )

        except Exception as e:
            self._log_warning(f"Failed to log existing content: {e}")

    def _verify_sheet_preservation(self, sheet_name: str) -> None:
        """Verify that non-updated sheet content is preserved."""
        try:
            sheet = self.workbook[sheet_name]

            # Count current content
            images = (
                len(sheet._images) if hasattr(sheet, "_images") and sheet._images else 0
            )
            charts = (
                len(sheet._charts) if hasattr(sheet, "_charts") and sheet._charts else 0
            )

            # Count non-empty cells
            data_cells = 0
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is not None and str(cell.value) != "!#VALUE":
                        data_cells += 1

            self._log_info(
                f"Verified preservation of sheet '{sheet_name}': {images} images, {charts} charts, {data_cells} data cells"
            )

            # Log any #VALUE errors found (indicating potential data loss)
            value_errors = 0
            for row in sheet.iter_rows():
                for cell in row:
                    if str(cell.value) == "!#VALUE":
                        value_errors += 1
                        self._log_warning(
                            f"Found !#VALUE error in {sheet_name}:{cell.coordinate}"
                        )

            if value_errors > 0:
                self._log_error(
                    f"Sheet '{sheet_name}' has {value_errors} !#VALUE errors - possible data loss"
                )

        except Exception as e:
            self._log_warning(
                f"Failed to verify sheet preservation for '{sheet_name}': {e}"
            )

    def _cleanup_temp_files(self) -> None:
        """Clean up temporary image files."""
        for temp_file in self.temp_image_files:
            try:
                if os.path.exists(temp_file):
                    os.unlink(temp_file)
            except Exception as e:
                logger.warning(f"Failed to cleanup temp file {temp_file}: {e}")
        self.temp_image_files.clear()

    def close(self) -> None:
        """Close workbook and cleanup resources."""
        self._cleanup_temp_files()
        if hasattr(self, "workbook"):
            self.workbook.close()

    def _insert_image_v2(
        self, sheet: Worksheet, cell_address: str, image_data: Any
    ) -> bool:
        """Insert image using direct file approach to avoid PIL corruption."""
        try:
            if not image_data:
                sheet[cell_address].value = "[NO_IMAGE_DATA]"
                self._log_warning(f"No image data provided for {cell_address}")
                return True

            # Determine final image path - avoid PIL processing that causes corruption
            final_image_path = None
            original_size = None

            if isinstance(image_data, str):
                if image_data.startswith("file://") or os.path.exists(image_data):
                    # Local file path - use directly when possible
                    try:
                        file_path = image_data
                        if file_path.startswith("file://"):
                            import urllib.parse

                            file_path = urllib.parse.unquote(file_path[7:])

                        if os.path.exists(file_path):
                            # Get original size without processing
                            with open(file_path, "rb") as f:
                                image_bytes = f.read()
                            temp_img = PILImage.open(io.BytesIO(image_bytes))
                            original_size = temp_img.size
                            temp_img.close()

                            # Use original file directly - no processing to avoid corruption
                            final_image_path = file_path
                            self._log_info(
                                f"Using original file directly: {file_path} (size: {original_size[0]}x{original_size[1]})"
                            )
                        else:
                            raise ValueError(f"File not found: {file_path}")
                    except Exception as e:
                        self._log_error(
                            f"Failed to process local file {image_data}: {e}"
                        )
                        sheet[cell_address].value = "!#VALUE"
                        return False

                elif image_data.startswith("data:image") or image_data.startswith(
                    "http"
                ):
                    # For base64 and URLs, save to temp file without resizing
                    try:
                        if image_data.startswith("data:image"):
                            image_bytes = self._decode_base64_image(image_data)
                        else:
                            image_bytes = self._download_image(image_data)

                        # Get size info
                        temp_img = PILImage.open(io.BytesIO(image_bytes))
                        original_size = temp_img.size
                        temp_img.close()

                        # Save raw bytes to temp file (no PIL processing)
                        import tempfile

                        temp_fd, temp_image_path = tempfile.mkstemp(suffix=".png")
                        os.close(temp_fd)
                        with open(temp_image_path, "wb") as f:
                            f.write(image_bytes)

                        self.temp_image_files.append(temp_image_path)
                        final_image_path = temp_image_path
                        self._log_info(
                            f"Saved to temp file: {temp_image_path} (size: {original_size[0]}x{original_size[1]})"
                        )
                    except Exception as e:
                        self._log_error(f"Failed to process image data: {e}")
                        sheet[cell_address].value = "!#VALUE"
                        return False
                else:
                    # Text fallback
                    sheet[cell_address].value = str(image_data)
                    self._log_info(
                        f"Inserted text instead of image at {cell_address}: {image_data}"
                    )
                    return True
            else:
                # Unknown format - use text
                sheet[cell_address].value = str(image_data)
                self._log_warning(
                    f"Unknown image format, inserted as text at {cell_address}"
                )
                return True

            if not final_image_path or not original_size:
                sheet[cell_address].value = "!#VALUE"
                self._log_error(f"Could not process image data for {cell_address}")
                return False

            # Create Excel image from final path (avoiding PIL altogether)
            excel_img = ExcelImage(final_image_path)

            # Apply width-constrained resizing with aspect ratio preservation
            target_width_points = 200  # Target width in points
            target_width_pixels = target_width_points * 1.33  # Convert to pixels

            # Calculate proper height based on original aspect ratio
            original_width, original_height = original_size
            aspect_ratio = original_height / original_width
            target_height_pixels = target_width_pixels * aspect_ratio

            # Set both width and height explicitly to maintain aspect ratio
            excel_img.width = target_width_pixels
            excel_img.height = target_height_pixels

            # Position at target cell
            row, col = self._parse_cell_address(cell_address)
            excel_img.anchor = f"{get_column_letter(col)}{row}"

            sheet.add_image(excel_img)
            # Explicitly ensure cell remains empty - no text pollution
            sheet[cell_address].value = None

            # Adjust row height based on new calculated height
            self._adjust_row_height(sheet, row, target_height_pixels)

            self._log_success(
                f"Inserted image at {cell_address} (original: {original_size[0]}x{original_size[1]}, resized: {target_width_pixels:.0f}x{target_height_pixels:.0f}px)"
            )
            return True

        except Exception as e:
            self._log_error(f"Image insert failed at {cell_address}: {e}")
            try:
                # Try to set error value if valid cell address
                if self._is_cell_address(cell_address):
                    sheet[cell_address].value = "!#VALUE"
            except:
                pass  # If even error setting fails, just continue
            return False
