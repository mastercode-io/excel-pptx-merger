"""Configuration schema validation for range images feature."""

import logging
from typing import Dict, Any, List, Optional
from dataclasses import dataclass

from .utils.exceptions import ValidationError

logger = logging.getLogger(__name__)


@dataclass
class RangeImageSchema:
    """Schema definition for range image configuration."""
    field_name: str
    sheet_name: str
    range: str
    include_headers: bool = True
    output_format: str = "png"
    dpi: int = 150
    fit_to_content: bool = True
    width: Optional[int] = None
    height: Optional[int] = None


class ConfigSchemaValidator:
    """Validates configuration schemas including range images."""
    
    SUPPORTED_VERSION = "1.1"
    VALID_IMAGE_FORMATS = ["png", "jpg", "jpeg"]
    MIN_DPI = 72
    MAX_DPI = 600
    MAX_RANGE_CELLS = 10000
    
    def __init__(self):
        """Initialize the schema validator."""
        self.validation_errors = []
        
    def validate_config(self, config: Dict[str, Any]) -> bool:
        """Validate the entire configuration including range images."""
        self.validation_errors = []
        
        try:
            # Validate basic structure
            self._validate_basic_structure(config)
            
            # Validate range images if present
            if "range_images" in config:
                self._validate_range_images(config["range_images"])
                
            # Validate global settings
            if "global_settings" in config:
                self._validate_global_settings(config["global_settings"])
                
            # Cross-validate range images with sheet configs
            if "range_images" in config and "sheet_configs" in config:
                self._cross_validate_range_images(
                    config["range_images"], 
                    config["sheet_configs"]
                )
                
            return len(self.validation_errors) == 0
            
        except Exception as e:
            self.validation_errors.append(f"Validation failed with error: {e}")
            return False
    
    def get_validation_errors(self) -> List[str]:
        """Get list of validation errors."""
        return self.validation_errors.copy()
    
    def _validate_basic_structure(self, config: Dict[str, Any]) -> None:
        """Validate basic configuration structure."""
        required_fields = ["version", "sheet_configs"]
        
        for field in required_fields:
            if field not in config:
                self.validation_errors.append(f"Missing required field: {field}")
                
        # Validate version
        if "version" in config:
            version = config["version"]
            if not isinstance(version, str):
                self.validation_errors.append("Version must be a string")
            elif version > self.SUPPORTED_VERSION:
                self.validation_errors.append(
                    f"Unsupported version: {version}. Max supported: {self.SUPPORTED_VERSION}"
                )
    
    def _validate_range_images(self, range_images: List[Dict[str, Any]]) -> None:
        """Validate range images configuration."""
        if not isinstance(range_images, list):
            self.validation_errors.append("range_images must be a list")
            return
            
        field_names = set()
        
        for i, range_config in enumerate(range_images):
            prefix = f"range_images[{i}]"
            
            # Validate required fields
            required_fields = ["field_name", "sheet_name", "range"]
            for field in required_fields:
                if field not in range_config:
                    self.validation_errors.append(f"{prefix}: Missing required field '{field}'")
                    
            # Validate field_name uniqueness
            field_name = range_config.get("field_name", "")
            if field_name in field_names:
                self.validation_errors.append(f"{prefix}: Duplicate field_name '{field_name}'")
            else:
                field_names.add(field_name)
                
            # Validate specific fields
            self._validate_range_config_fields(range_config, prefix)
    
    def _validate_range_config_fields(self, config: Dict[str, Any], prefix: str) -> None:
        """Validate individual range configuration fields."""
        # Validate range format
        if "range" in config:
            if not self._is_valid_excel_range(config["range"]):
                self.validation_errors.append(f"{prefix}: Invalid range format '{config['range']}'")
            else:
                # Check range size
                if self._get_range_cell_count(config["range"]) > self.MAX_RANGE_CELLS:
                    self.validation_errors.append(
                        f"{prefix}: Range too large (max {self.MAX_RANGE_CELLS} cells)"
                    )
        
        # Validate output_format
        if "output_format" in config:
            format_val = config["output_format"].lower()
            if format_val not in self.VALID_IMAGE_FORMATS:
                self.validation_errors.append(
                    f"{prefix}: Invalid output_format '{format_val}'. "
                    f"Valid: {self.VALID_IMAGE_FORMATS}"
                )
        
        # Validate DPI
        if "dpi" in config:
            dpi = config["dpi"]
            if not isinstance(dpi, int) or dpi < self.MIN_DPI or dpi > self.MAX_DPI:
                self.validation_errors.append(
                    f"{prefix}: DPI must be integer between {self.MIN_DPI} and {self.MAX_DPI}"
                )
        
        # Validate dimensions
        for dim in ["width", "height"]:
            if dim in config:
                value = config[dim]
                if value is not None and (not isinstance(value, int) or value <= 0):
                    self.validation_errors.append(f"{prefix}: {dim} must be positive integer or null")
        
        # Validate boolean fields
        boolean_fields = ["include_headers", "fit_to_content"]
        for field in boolean_fields:
            if field in config and not isinstance(config[field], bool):
                self.validation_errors.append(f"{prefix}: {field} must be boolean")
    
    def _validate_global_settings(self, global_settings: Dict[str, Any]) -> None:
        """Validate global settings including range image settings."""
        if "range_images" in global_settings:
            range_settings = global_settings["range_images"]
            
            if not isinstance(range_settings, dict):
                self.validation_errors.append("global_settings.range_images must be object")
                return
            
            # Validate enabled flag
            if "enabled" in range_settings:
                if not isinstance(range_settings["enabled"], bool):
                    self.validation_errors.append("range_images.enabled must be boolean")
            
            # Validate max_range_cells
            if "max_range_cells" in range_settings:
                max_cells = range_settings["max_range_cells"]
                if not isinstance(max_cells, int) or max_cells <= 0:
                    self.validation_errors.append("range_images.max_range_cells must be positive integer")
            
            # Validate default_dpi
            if "default_dpi" in range_settings:
                dpi = range_settings["default_dpi"]
                if not isinstance(dpi, int) or dpi < self.MIN_DPI or dpi > self.MAX_DPI:
                    self.validation_errors.append(
                        f"range_images.default_dpi must be between {self.MIN_DPI} and {self.MAX_DPI}"
                    )
            
            # Validate default_format
            if "default_format" in range_settings:
                format_val = range_settings["default_format"].lower()
                if format_val not in self.VALID_IMAGE_FORMATS:
                    self.validation_errors.append(
                        f"range_images.default_format must be one of: {self.VALID_IMAGE_FORMATS}"
                    )
    
    def _cross_validate_range_images(self, range_images: List[Dict[str, Any]], 
                                   sheet_configs: Dict[str, Any]) -> None:
        """Cross-validate range images against sheet configurations."""
        available_sheets = set(sheet_configs.keys())
        
        for i, range_config in enumerate(range_images):
            prefix = f"range_images[{i}]"
            sheet_name = range_config.get("sheet_name", "")
            
            if sheet_name and sheet_name not in available_sheets:
                self.validation_errors.append(
                    f"{prefix}: Sheet '{sheet_name}' not found in sheet_configs. "
                    f"Available: {list(available_sheets)}"
                )
    
    def _is_valid_excel_range(self, range_str: str) -> bool:
        """Check if range string has valid Excel format."""
        import re
        # Pattern for Excel range: Letter(s) + Number + : + Letter(s) + Number
        pattern = r'^[A-Z]+\d+:[A-Z]+\d+$'
        return bool(re.match(pattern, range_str.upper()))
    
    def _get_range_cell_count(self, range_str: str) -> int:
        """Calculate number of cells in a range."""
        try:
            # Parse range like "A1:C10"
            parts = range_str.upper().split(":")
            if len(parts) != 2:
                return 0
                
            start_cell, end_cell = parts
            
            # Extract column letters and row numbers
            import re
            start_match = re.match(r'^([A-Z]+)(\d+)$', start_cell)
            end_match = re.match(r'^([A-Z]+)(\d+)$', end_cell)
            
            if not start_match or not end_match:
                return 0
                
            start_col, start_row = start_match.groups()
            end_col, end_row = end_match.groups()
            
            # Convert column letters to numbers
            start_col_num = self._column_letters_to_number(start_col)
            end_col_num = self._column_letters_to_number(end_col)
            
            # Calculate cell count
            cols = end_col_num - start_col_num + 1
            rows = int(end_row) - int(start_row) + 1
            
            return cols * rows
            
        except Exception:
            return 0
    
    def _column_letters_to_number(self, letters: str) -> int:
        """Convert Excel column letters to number (A=1, B=2, etc.)."""
        result = 0
        for char in letters:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result


def validate_config_file(config: Dict[str, Any]) -> tuple[bool, List[str]]:
    """Validate configuration file and return validation result."""
    validator = ConfigSchemaValidator()
    is_valid = validator.validate_config(config)
    errors = validator.get_validation_errors()
    
    if errors:
        logger.warning(f"Configuration validation found {len(errors)} errors")
        for error in errors:
            logger.warning(f"  - {error}")
    else:
        logger.info("Configuration validation passed")
    
    return is_valid, errors


def create_range_image_schema(config_dict: Dict[str, Any]) -> RangeImageSchema:
    """Create RangeImageSchema from configuration dictionary."""
    try:
        return RangeImageSchema(
            field_name=config_dict["field_name"],
            sheet_name=config_dict["sheet_name"],
            range=config_dict["range"],
            include_headers=config_dict.get("include_headers", True),
            output_format=config_dict.get("output_format", "png"),
            dpi=config_dict.get("dpi", 150),
            fit_to_content=config_dict.get("fit_to_content", True),
            width=config_dict.get("width"),
            height=config_dict.get("height")
        )
    except KeyError as e:
        raise ValidationError(f"Missing required field in range image config: {e}")
    except Exception as e:
        raise ValidationError(f"Invalid range image configuration: {e}")