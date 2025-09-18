"""Input validation utilities for Excel to PowerPoint Merger."""

import re
from typing import Any, Dict, List, Optional, Union
import jsonschema
from jsonschema import validate, ValidationError as JsonSchemaValidationError

from .exceptions import ValidationError


def validate_json_schema(data: Dict[str, Any], schema: Dict[str, Any]) -> None:
    """Validate data against JSON schema."""
    try:
        validate(instance=data, schema=schema)
    except JsonSchemaValidationError as e:
        raise ValidationError(f"Schema validation failed: {e.message}")


def validate_config_structure(config: Dict[str, Any]) -> None:
    """Validate configuration structure."""
    schema = {
        "type": "object",
        "properties": {
            "version": {"type": "string"},
            "sheet_configs": {
                "type": "object",
                "patternProperties": {
                    ".*": {
                        "type": "object",
                        "properties": {
                            "subtables": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "name": {"type": "string"},
                                        "type": {
                                            "type": "string",
                                            "enum": ["key_value_pairs", "table"],
                                        },
                                        "header_search": {
                                            "type": "object",
                                            "properties": {
                                                "method": {"type": "string"},
                                                "text": {"type": "string"},
                                                "column": {"type": "string"},
                                                "search_range": {"type": "string"},
                                            },
                                            "required": ["method"],
                                        },
                                        "data_extraction": {
                                            "type": "object",
                                            "properties": {
                                                "orientation": {
                                                    "type": "string",
                                                    "enum": ["horizontal", "vertical"],
                                                },
                                                "headers_row_offset": {
                                                    "type": "integer"
                                                },
                                                "headers_col_offset": {
                                                    "type": "integer"
                                                },
                                                "data_row_offset": {"type": "integer"},
                                                "max_columns": {"type": "integer"},
                                                "max_rows": {"type": "integer"},
                                                "column_mappings": {
                                                    "type": "object",
                                                    "patternProperties": {
                                                        ".*": {
                                                            "oneOf": [
                                                                {
                                                                    "type": "string"  # For backward compatibility
                                                                },
                                                                {
                                                                    "type": "object",
                                                                    "properties": {
                                                                        "name": {
                                                                            "type": "string"
                                                                        },
                                                                        "type": {
                                                                            "type": "string",
                                                                            "enum": [
                                                                                "text",
                                                                                "image",
                                                                                "number",
                                                                                "date",
                                                                                "boolean",
                                                                                "link",
                                                                            ],
                                                                        },
                                                                    },
                                                                    "required": [
                                                                        "name"
                                                                    ],
                                                                },
                                                            ]
                                                        }
                                                    },
                                                },
                                            },
                                        },
                                        "data_update": {
                                            "type": "object",
                                            "properties": {
                                                "orientation": {
                                                    "type": "string",
                                                    "enum": ["horizontal", "vertical"],
                                                },
                                                "headers_row_offset": {
                                                    "type": "integer"
                                                },
                                                "headers_col_offset": {
                                                    "type": "integer"
                                                },
                                                "data_row_offset": {"type": "integer"},
                                                "max_columns": {"type": "integer"},
                                                "max_rows": {"type": "integer"},
                                                "column_mappings": {
                                                    "type": "object",
                                                    "patternProperties": {
                                                        ".*": {
                                                            "type": "object",
                                                            "properties": {
                                                                "name": {
                                                                    "type": "string"
                                                                },
                                                                "type": {
                                                                    "type": "string",
                                                                    "enum": [
                                                                        "text",
                                                                        "image",
                                                                        "number",
                                                                        "date",
                                                                        "boolean",
                                                                        "link",
                                                                    ],
                                                                },
                                                            },
                                                            "required": [
                                                                "name",
                                                                "type",
                                                            ],
                                                        }
                                                    },
                                                },
                                            },
                                            "required": ["column_mappings"],
                                        },
                                    },
                                    "anyOf": [
                                        {
                                            "required": [
                                                "name",
                                                "type",
                                                "header_search",
                                                "data_extraction",
                                            ]
                                        },
                                        {
                                            "required": [
                                                "name",
                                                "type",
                                                "header_search",
                                                "data_update",
                                            ]
                                        },
                                    ],
                                },
                            }
                        },
                        "required": ["subtables"],
                    }
                },
            },
            "global_settings": {
                "type": "object",
                "properties": {
                    "normalize_keys": {"type": "boolean"},
                    "temp_file_cleanup": {
                        "type": "object",
                        "properties": {
                            "enabled": {"type": "boolean"},
                            "delay_seconds": {"type": "integer"},
                            "keep_on_error": {"type": "boolean"},
                            "development_mode": {"type": "boolean"},
                        },
                    },
                    "image_storage": {
                        "type": "object",
                        "properties": {
                            "development_mode": {
                                "type": "object",
                                "properties": {
                                    "directory": {"type": "string"},
                                    "cleanup_after_merge": {"type": "boolean"},
                                },
                            },
                            "production_mode": {
                                "type": "object",
                                "properties": {
                                    "directory": {"type": "string"},
                                    "cleanup_after_merge": {"type": "boolean"},
                                },
                            },
                        },
                    },
                },
            },
        },
        "required": ["version", "sheet_configs"],
    }

    validate_json_schema(config, schema)


def get_field_type_from_mapping(column_mapping: Union[str, Dict[str, Any]]) -> str:
    """Extract field type from column mapping.

    Supports both old format (string) and new format (object with name and type).
    Default type is 'text' if not specified.
    """
    if isinstance(column_mapping, str):
        return "text"  # Default type for backward compatibility

    return column_mapping.get("type", "text")


def validate_merge_fields(template_text: str) -> List[str]:
    """Extract and validate merge fields from template text."""
    merge_field_pattern = r"\{\{([^}]+)\}\}"
    fields = re.findall(merge_field_pattern, template_text)

    validated_fields = []
    for field in fields:
        field = field.strip()
        if field:
            # Validate field name format
            if not re.match(r"^[a-zA-Z_][a-zA-Z0-9_]*(\.[a-zA-Z0-9_]+)*$", field):
                # Log warning for invalid field but continue processing
                import logging

                logger = logging.getLogger(__name__)
                logger.warning(f"Skipping invalid merge field format: {{{{{field}}}}}")
                continue
            validated_fields.append(field)

    return validated_fields


def clean_excel_text_value(value: Any, clean_quotes: bool = True) -> Any:
    """Clean Excel cell values by removing leading single quotes.
    
    Excel users often prefix text with a single quote (') to force text formatting
    and prevent automatic type conversion. This function removes that leading quote
    while preserving the quote if it's part of the actual content.
    
    Args:
        value: The cell value to clean
        clean_quotes: Whether to perform quote cleaning (default: True)
        
    Returns:
        Cleaned value with leading quote removed if applicable
    """
    if not clean_quotes or value is None:
        return value
        
    # Only process string values
    if not isinstance(value, str):
        return value
        
    # Only remove leading single quote, preserve quotes elsewhere
    if value.startswith("'") and len(value) > 1:
        return value[1:]
        
    return value


def get_clean_quotes_setting(config: Dict[str, Any]) -> bool:
    """Get the clean_excel_quotes setting from configuration.
    
    Args:
        config: Configuration dictionary
        
    Returns:
        Boolean indicating whether to clean Excel quotes (default: True)
    """
    return config.get("global_settings", {}).get("clean_excel_quotes", True)


def normalize_column_name(column_name: str) -> str:
    """Normalize Excel column names to valid JSON keys."""
    if not column_name or not isinstance(column_name, str):
        return "unnamed_column"

    # Convert to lowercase
    normalized = column_name.lower().strip()

    # Replace spaces and special characters with underscores
    normalized = re.sub(r"[^\w]+", "_", normalized)

    # Remove multiple underscores
    normalized = re.sub(r"_+", "_", normalized)

    # Remove leading/trailing underscores
    normalized = normalized.strip("_")

    # Ensure it starts with a letter or underscore
    if normalized and not re.match(r"^[a-zA-Z_]", normalized):
        normalized = f"col_{normalized}"

    return normalized or "unnamed_column"


def validate_cell_range(cell_range: str) -> bool:
    """Validate Excel cell range format."""
    if not cell_range:
        return False

    # Pattern for cell range like A1:B10 or A1
    pattern = r"^[A-Z]+\d+(:[A-Z]+\d+)?$"
    return bool(re.match(pattern, cell_range.upper()))


def validate_column_reference(column_ref: str) -> bool:
    """Validate Excel column reference format."""
    if not column_ref:
        return False

    # Pattern for column reference like A, B, AA, etc.
    pattern = r"^[A-Z]+$"
    return bool(re.match(pattern, column_ref.upper()))


def sanitize_filename(filename: str) -> str:
    """Sanitize filename for safe file system usage."""
    if not filename:
        return "unnamed_file"

    # Remove or replace dangerous characters
    sanitized = re.sub(r'[<>:"/\\|?*]', "_", filename)

    # Remove control characters
    sanitized = re.sub(r"[\x00-\x1f\x7f-\x9f]", "", sanitized)

    # Limit length
    sanitized = sanitized[:255]

    # Ensure not empty
    return sanitized or "unnamed_file"


def validate_data_type(value: Any, expected_type: type) -> bool:
    """Validate if value matches expected type."""
    try:
        if expected_type == str:
            return isinstance(value, (str, int, float)) or value is None
        elif expected_type == int:
            return isinstance(value, (int, float)) and not isinstance(value, bool)
        elif expected_type == float:
            return isinstance(value, (int, float)) and not isinstance(value, bool)
        elif expected_type == bool:
            return isinstance(value, bool)
        else:
            return isinstance(value, expected_type)
    except Exception:
        return False


def validate_api_request(request_data: Dict[str, Any]) -> None:
    """Validate API request data structure."""
    if not isinstance(request_data, dict):
        raise ValidationError("Request data must be a JSON object")

    # Check for required files
    if "excel_file" not in request_data and "pptx_file" not in request_data:
        raise ValidationError("Both excel_file and pptx_file are required")

    # Validate configuration if provided
    if "config" in request_data:
        if not isinstance(request_data["config"], dict):
            raise ValidationError("Configuration must be a JSON object")


def is_empty_cell_value(value: Any) -> bool:
    """Check if cell value should be considered empty."""
    if value is None:
        return True

    if isinstance(value, str):
        return not value.strip()

    if isinstance(value, (int, float)):
        return False  # Numbers are never considered empty

    return not bool(value)
