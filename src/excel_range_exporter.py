"""Excel range to image exporter using Microsoft Graph API."""

import logging
import os
import tempfile
import time
from typing import Dict, Any, Optional, List, Tuple
from dataclasses import dataclass
from pathlib import Path

from .graph_api_client import GraphAPIClient, GraphAPIError
from .temp_file_manager import TempFileManager
from .utils.exceptions import ExcelProcessingError, ValidationError

logger = logging.getLogger(__name__)


@dataclass
class RangeImageConfig:
    """Configuration for range image export."""

    field_name: str
    sheet_name: str
    range: str
    include_headers: bool = True
    output_format: str = "png"
    dpi: int = 150
    fit_to_content: bool = True
    width: Optional[int] = None
    height: Optional[int] = None


@dataclass
class RangeImageResult:
    """Result of range image export."""

    field_name: str
    image_path: str
    image_data: bytes
    width: int
    height: int
    range_dimensions: Tuple[int, int]  # (rows, columns)
    success: bool
    error_message: Optional[str] = None


class ExcelRangeExporter:
    """Exports Excel ranges as images using Microsoft Graph API."""

    def __init__(self, client_id: str, client_secret: str, tenant_id: str):
        """Initialize with Azure app credentials."""
        self.graph_client = GraphAPIClient(client_id, client_secret, tenant_id)
        self.temp_manager = TempFileManager()

    def export_ranges(
        self, workbook_path: str, range_configs: List[RangeImageConfig]
    ) -> List[RangeImageResult]:
        """Export multiple ranges from an Excel workbook."""
        if not os.path.exists(workbook_path):
            raise ExcelProcessingError(f"Workbook not found: {workbook_path}")

        if not range_configs:
            logger.warning("No range configurations provided")
            return []

        results = []
        item_id = None

        try:
            # Upload workbook to OneDrive
            logger.info(f"Uploading workbook to OneDrive: {workbook_path}")
            item_id = self.graph_client.upload_workbook_to_onedrive(workbook_path)

            # Get available worksheets for validation
            worksheets = self.graph_client.get_worksheet_names(item_id)
            logger.info(f"Available worksheets: {worksheets}")

            # Process each range configuration
            for config in range_configs:
                result = self._export_single_range(item_id, config, worksheets)
                results.append(result)

        except GraphAPIError as e:
            logger.error(f"Graph API error during range export: {e}")
            raise ExcelProcessingError(f"Failed to export ranges: {e}")
        except Exception as e:
            logger.error(f"Unexpected error during range export: {e}")
            raise ExcelProcessingError(f"Unexpected error: {e}")
        finally:
            # Cleanup temporary file from OneDrive
            if item_id:
                try:
                    self.graph_client.cleanup_temp_file(item_id)
                except Exception as e:
                    logger.warning(f"Failed to cleanup OneDrive file: {e}")

        return results

    def _export_single_range(
        self, item_id: str, config: RangeImageConfig, available_worksheets: List[str]
    ) -> RangeImageResult:
        """Export a single range as image."""
        try:
            # Validate sheet exists
            if config.sheet_name not in available_worksheets:
                error_msg = f"Sheet '{config.sheet_name}' not found. Available: {available_worksheets}"
                logger.error(error_msg)
                return RangeImageResult(
                    field_name=config.field_name,
                    image_path="",
                    image_data=b"",
                    width=0,
                    height=0,
                    range_dimensions=(0, 0),
                    success=False,
                    error_message=error_msg,
                )

            # Validate range exists and has data
            if not self.graph_client.validate_range(
                item_id, config.sheet_name, config.range
            ):
                error_msg = f"Range '{config.range}' in sheet '{config.sheet_name}' is invalid or empty"
                logger.warning(error_msg)
                return RangeImageResult(
                    field_name=config.field_name,
                    image_path="",
                    image_data=b"",
                    width=0,
                    height=0,
                    range_dimensions=(0, 0),
                    success=False,
                    error_message=error_msg,
                )

            # Get range dimensions
            range_dimensions = self.graph_client.get_range_dimensions(
                item_id, config.sheet_name, config.range
            )

            # Render range as image
            logger.info(
                f"Rendering range {config.range} from sheet '{config.sheet_name}' as image"
            )
            image_data = self.graph_client.render_range_as_image(
                item_id=item_id,
                sheet_name=config.sheet_name,
                range_str=config.range,
                width=config.width,
                height=config.height,
            )

            # Save image to temporary file
            image_path = self._save_image_to_temp_file(
                image_data, config.field_name, config.output_format
            )

            # Get actual image dimensions
            actual_width, actual_height = self._get_image_dimensions(image_data)

            logger.info(f"Successfully exported range {config.range} to {image_path}")
            return RangeImageResult(
                field_name=config.field_name,
                image_path=image_path,
                image_data=image_data,
                width=actual_width,
                height=actual_height,
                range_dimensions=range_dimensions,
                success=True,
            )

        except GraphAPIError as e:
            error_msg = f"Failed to export range {config.range}: {e}"
            logger.error(error_msg)
            return RangeImageResult(
                field_name=config.field_name,
                image_path="",
                image_data=b"",
                width=0,
                height=0,
                range_dimensions=(0, 0),
                success=False,
                error_message=error_msg,
            )
        except Exception as e:
            error_msg = f"Unexpected error exporting range {config.range}: {e}"
            logger.error(error_msg)
            return RangeImageResult(
                field_name=config.field_name,
                image_path="",
                image_data=b"",
                width=0,
                height=0,
                range_dimensions=(0, 0),
                success=False,
                error_message=error_msg,
            )

    def _save_image_to_temp_file(
        self, image_data: bytes, field_name: str, output_format: str
    ) -> str:
        """Save image data to temporary file."""
        # Generate unique filename
        timestamp = int(time.time())
        filename = f"range_image_{field_name}_{timestamp}.{output_format.lower()}"

        # Use temp file manager for consistent handling
        temp_path = self.temp_manager.create_temp_file(
            suffix=f".{output_format.lower()}"
        )

        try:
            with open(temp_path, "wb") as f:
                f.write(image_data)

            logger.debug(f"Saved range image to temporary file: {temp_path}")
            return temp_path

        except Exception as e:
            logger.error(f"Failed to save image to temporary file: {e}")
            raise ExcelProcessingError(f"Failed to save image: {e}")

    def _get_image_dimensions(self, image_data: bytes) -> Tuple[int, int]:
        """Get image dimensions from binary data."""
        try:
            from PIL import Image
            import io

            image = Image.open(io.BytesIO(image_data))
            return image.size  # (width, height)

        except Exception as e:
            logger.warning(f"Failed to get image dimensions: {e}")
            return (0, 0)

    def validate_config(self, config: RangeImageConfig) -> bool:
        """Validate a range image configuration."""
        try:
            # Check required fields
            if not config.field_name:
                raise ValidationError("field_name is required")
            if not config.sheet_name:
                raise ValidationError("sheet_name is required")
            if not config.range:
                raise ValidationError("range is required")

            # Validate range format (basic check)
            if not self._is_valid_range_format(config.range):
                raise ValidationError(f"Invalid range format: {config.range}")

            # Validate output format
            valid_formats = ["png", "jpg", "jpeg"]
            if config.output_format.lower() not in valid_formats:
                raise ValidationError(f"Invalid output format: {config.output_format}")

            # Validate DPI
            if config.dpi < 72 or config.dpi > 600:
                raise ValidationError(
                    f"DPI must be between 72 and 600, got: {config.dpi}"
                )

            return True

        except ValidationError:
            raise
        except Exception as e:
            raise ValidationError(f"Configuration validation failed: {e}")

    def _is_valid_range_format(self, range_str: str) -> bool:
        """Check if range string has valid Excel format (e.g., A1:E15)."""
        import re

        # Basic pattern for Excel range: Letter(s) + Number + : + Letter(s) + Number
        pattern = r"^[A-Z]+\d+:[A-Z]+\d+$"
        return bool(re.match(pattern, range_str.upper()))

    def cleanup_temp_files(self) -> None:
        """Clean up temporary files created during export."""
        try:
            self.temp_manager.cleanup_all()
            logger.info("Cleaned up temporary range image files")
        except Exception as e:
            logger.warning(f"Failed to cleanup temporary files: {e}")


def create_range_configs_from_dict(
    range_configs_data: List[Dict[str, Any]],
) -> List[RangeImageConfig]:
    """Create RangeImageConfig objects from dictionary data."""
    configs = []

    for config_data in range_configs_data:
        try:
            config = RangeImageConfig(**config_data)
            configs.append(config)
        except TypeError as e:
            logger.error(f"Invalid range configuration: {config_data}, error: {e}")
            raise ValidationError(f"Invalid range configuration: {e}")

    return configs


def validate_range_configs(configs: List[RangeImageConfig]) -> List[str]:
    """Validate multiple range configurations and return list of errors."""
    errors = []
    field_names = set()

    for i, config in enumerate(configs):
        try:
            # Check for duplicate field names
            if config.field_name in field_names:
                errors.append(f"Config {i}: Duplicate field_name '{config.field_name}'")
            else:
                field_names.add(config.field_name)

            # Validate individual config (this will raise ValidationError if invalid)
            exporter = ExcelRangeExporter("", "", "")  # Dummy exporter for validation
            exporter.validate_config(config)

        except ValidationError as e:
            errors.append(f"Config {i}: {e}")
        except Exception as e:
            errors.append(f"Config {i}: Unexpected error - {e}")

    return errors
