"""Request handler utilities for dual-mode payload processing.

This module provides utilities for handling both multipart/form-data and application/json
payloads across all API endpoints, ensuring consistent processing and CRM compatibility.
"""

import base64
import io
import json
import logging
from typing import Dict, Any, Optional, Tuple, Union, BinaryIO
from werkzeug.datastructures import FileStorage
from flask import Request

from src.utils.exceptions import ValidationError

logger = logging.getLogger(__name__)


class RequestPayloadDetector:
    """Detects and analyzes request payload type (JSON vs multipart)."""

    @staticmethod
    def detect_payload_mode(request: Request) -> Tuple[bool, bool, bool]:
        """Detect the payload mode of the request.

        Args:
            request: Flask request object

        Returns:
            Tuple of (is_json_request, has_form_data, has_files)
        """
        content_type = request.headers.get("Content-Type", "Not specified")
        is_json_request = request.is_json
        has_form_data = bool(request.form)
        has_files = bool(request.files)

        # Fallback JSON detection for CRM systems that send JSON with wrong Content-Type
        if not is_json_request and request.data and not has_form_data and not has_files:
            try:
                # Try to parse raw request data as JSON
                json.loads(request.data)
                is_json_request = True
                logger.info(
                    f"Detected JSON payload despite Content-Type: {content_type}"
                )
            except json.JSONDecodeError:
                logger.debug("Raw request data is not valid JSON")

        logger.info(
            f"Request analysis - JSON: {is_json_request}, Form data: {has_form_data}, Files: {has_files}"
        )

        return is_json_request, has_form_data, has_files

    @staticmethod
    def log_request_info(request: Request) -> None:
        """Log detailed request information for debugging."""
        logger.info("📊 Request method: %s", request.method)
        logger.info("📊 Request path: %s", request.path)
        logger.info("📊 Has files: %s", bool(request.files))
        logger.info("📊 Has form: %s", bool(request.form))
        logger.info("📊 Is JSON: %s", request.is_json)
        logger.info(
            "📊 User-Agent: %s", request.headers.get("User-Agent", "Not specified")
        )
        logger.info(
            "📊 Content-Type: %s", request.headers.get("Content-Type", "Not specified")
        )
        logger.info(
            "📊 Content-Length: %s",
            request.headers.get("Content-Length", "Not specified"),
        )

        # Check for CRM system indicators
        user_agent = request.headers.get("User-Agent", "").lower()
        crm_indicators = ["deluge", "zoho", "crm", "automation", "webhook", "postman"]
        is_likely_crm = any(indicator in user_agent for indicator in crm_indicators)
        if is_likely_crm:
            logger.info("🏢 POTENTIAL CRM/AUTOMATION SYSTEM DETECTED!")


class PayloadParser:
    """Parses data from both JSON and multipart payloads."""

    def __init__(self, request: Request, is_json_request: bool):
        """Initialize the parser with request context.

        Args:
            request: Flask request object
            is_json_request: Whether this is a JSON request
        """
        self.request = request
        self.is_json_request = is_json_request
        self._json_data = None

    def get_json_data(self) -> Dict[str, Any]:
        """Get JSON data from the request, handling CRM compatibility.

        Returns:
            Parsed JSON data

        Raises:
            ValidationError: If JSON parsing fails
        """
        if not self.is_json_request:
            raise ValidationError("Not a JSON request")

        if self._json_data is not None:
            return self._json_data

        try:
            # Handle both standard JSON requests and CRM systems with wrong Content-Type
            if self.request.is_json:
                logger.info("📝 Attempting standard request.get_json()...")
                self._json_data = self.request.get_json()
                logger.info("✅ Standard JSON parsing succeeded")
            else:
                # Parse raw data for systems that send JSON with text/plain Content-Type
                logger.info(
                    "📝 Attempting json.loads(request.data) for CRM compatibility..."
                )
                self._json_data = json.loads(self.request.data)
                logger.info("✅ CRM compatibility JSON parsing succeeded")
                logger.info(
                    "Parsed JSON from raw request data due to incorrect Content-Type"
                )

            if not self._json_data:
                raise ValidationError("JSON payload is empty")

            return self._json_data

        except json.JSONDecodeError as e:
            logger.error("❌ JSON parsing failed: %s", e)
            raise ValidationError(f"Invalid JSON format: {e}")

    def get_file(
        self, field_name: str, required: bool = True
    ) -> Optional[Union[FileStorage, io.BytesIO]]:
        """Get file from either multipart or JSON (base64) payload.

        Args:
            field_name: Name of the file field
            required: Whether the file is required

        Returns:
            File object (FileStorage for multipart, BytesIO for JSON) or None

        Raises:
            ValidationError: If required file is missing or invalid
        """
        if self.is_json_request:
            # JSON mode: Extract base64-encoded file
            json_data = self.get_json_data()
            file_b64 = json_data.get(field_name)

            if not file_b64:
                if required:
                    raise ValidationError(
                        f"{field_name} (base64) is required in JSON mode"
                    )
                return None

            logger.info(f"Base64 {field_name} size: {len(file_b64)} characters")

            try:
                file_data = base64.b64decode(file_b64)
                file_obj = io.BytesIO(file_data)
                # Add filename attribute for compatibility
                file_obj.filename = json_data.get(
                    f"{field_name}_name", f"{field_name}.bin"
                )
                logger.info(f"Decoded {field_name} size: {len(file_data)} bytes")
                return file_obj
            except Exception as e:
                raise ValidationError(f"Invalid base64 {field_name}: {e}")
        else:
            # Multipart mode: Get file from request.files
            if field_name not in self.request.files:
                if required:
                    raise ValidationError(f"{field_name} is required")
                return None

            return self.request.files[field_name]

    def get_param(
        self, param_name: str, default: Any = None, required: bool = False
    ) -> Any:
        """Get parameter from either form data or JSON payload.

        Args:
            param_name: Name of the parameter
            default: Default value if not found
            required: Whether the parameter is required

        Returns:
            Parameter value

        Raises:
            ValidationError: If required parameter is missing
        """
        if self.is_json_request:
            json_data = self.get_json_data()
            value = json_data.get(param_name, default)
        else:
            value = self.request.form.get(param_name, default)

        if required and value is None:
            raise ValidationError(f"{param_name} parameter is required")

        return value

    def get_json_param(
        self, param_name: str, default: Any = None, required: bool = False
    ) -> Any:
        """Get JSON parameter (parses JSON string in multipart, direct object in JSON mode).

        Args:
            param_name: Name of the parameter
            default: Default value if not found
            required: Whether the parameter is required

        Returns:
            Parsed parameter value

        Raises:
            ValidationError: If required parameter is missing or invalid JSON
        """
        if self.is_json_request:
            # In JSON mode, parameters are already objects
            json_data = self.get_json_data()
            value = json_data.get(param_name, default)
        else:
            # In multipart mode, JSON parameters are strings that need parsing
            value_str = self.request.form.get(param_name)
            if not value_str:
                if required:
                    raise ValidationError(f"{param_name} parameter is required")
                return default

            try:
                value = json.loads(value_str)
            except json.JSONDecodeError as e:
                raise ValidationError(f"Invalid JSON in {param_name} parameter: {e}")

        if required and value is None:
            raise ValidationError(f"{param_name} parameter is required")

        return value

    def get_sharepoint_info(self) -> Tuple[Optional[str], Optional[str]]:
        """Get SharePoint URL and item ID from the request.

        This method supports both naming conventions for backward compatibility:
        - Legacy: sharepoint_file_url, sharepoint_item_id
        - Standard: sharepoint_excel_url, sharepoint_excel_id

        Returns:
            Tuple of (sharepoint_url, sharepoint_item_id)
        """
        # Check for type-specific parameters first (standard convention)
        sharepoint_url = self.get_param("sharepoint_excel_url")
        sharepoint_item_id = self.get_param("sharepoint_excel_id")

        # Fall back to generic parameters for backward compatibility
        if not sharepoint_url:
            sharepoint_url = self.get_param("sharepoint_file_url")
        if not sharepoint_item_id:
            sharepoint_item_id = self.get_param("sharepoint_item_id")

        return sharepoint_url, sharepoint_item_id

    def get_sharepoint_info_extended(self) -> Dict[str, Optional[str]]:
        """Get all SharePoint parameters from the request.

        Returns all possible SharePoint parameters for both Excel and PowerPoint files.

        Returns:
            Dict with keys: sharepoint_excel_url, sharepoint_excel_id,
                          sharepoint_pptx_url, sharepoint_pptx_id
        """
        return {
            "sharepoint_excel_url": self.get_param("sharepoint_excel_url"),
            "sharepoint_excel_id": self.get_param("sharepoint_excel_id"),
            "sharepoint_pptx_url": self.get_param("sharepoint_pptx_url"),
            "sharepoint_pptx_id": self.get_param("sharepoint_pptx_id"),
        }

    def get_sharepoint_file(
        self,
        sharepoint_config: Dict[str, Any],
        default_filename: str = "sharepoint_file.xlsx",
    ) -> Optional[io.BytesIO]:
        """Get file from SharePoint using centralized handler.

        Args:
            sharepoint_config: SharePoint configuration from global_settings
            default_filename: Default filename for downloaded file

        Returns:
            BytesIO object with downloaded file, or None if no SharePoint reference

        Raises:
            ValidationError: If SharePoint access fails
        """
        sharepoint_url, sharepoint_item_id = self.get_sharepoint_info()

        # Return None if no SharePoint reference
        if not sharepoint_url and not sharepoint_item_id:
            return None

        # Use centralized SharePoint handler
        from .sharepoint_file_handler import SharePointFileHandler

        sp_handler = SharePointFileHandler(sharepoint_config)
        return sp_handler.download_file(
            sharepoint_url=sharepoint_url,
            sharepoint_item_id=sharepoint_item_id,
            default_filename=default_filename,
        )

    def get_sharepoint_excel_file(
        self, sharepoint_config: Dict[str, Any]
    ) -> Optional[io.BytesIO]:
        """Get Excel file from SharePoint using excel-specific parameters."""
        sharepoint_url = self.get_param("sharepoint_excel_url")
        sharepoint_item_id = self.get_param("sharepoint_excel_id")

        if not sharepoint_url and not sharepoint_item_id:
            return None

        from .sharepoint_file_handler import SharePointFileHandler

        sp_handler = SharePointFileHandler(sharepoint_config)
        return sp_handler.download_file(
            sharepoint_url=sharepoint_url,
            sharepoint_item_id=sharepoint_item_id,
            default_filename="sharepoint_excel.xlsx",
        )

    def get_sharepoint_pptx_file(
        self, sharepoint_config: Dict[str, Any]
    ) -> Optional[io.BytesIO]:
        """Get PowerPoint file from SharePoint using pptx-specific parameters."""
        sharepoint_url = self.get_param("sharepoint_pptx_url")
        sharepoint_item_id = self.get_param("sharepoint_pptx_id")

        if not sharepoint_url and not sharepoint_item_id:
            return None

        from .sharepoint_file_handler import SharePointFileHandler

        sp_handler = SharePointFileHandler(sharepoint_config)
        return sp_handler.download_file(
            sharepoint_url=sharepoint_url,
            sharepoint_item_id=sharepoint_item_id,
            default_filename="sharepoint_template.pptx",
        )

    def get_file_data(self, file_obj: Union[FileStorage, io.BytesIO]) -> bytes:
        """Get raw bytes from a file object.

        Args:
            file_obj: File object (FileStorage or BytesIO)

        Returns:
            File data as bytes
        """
        if isinstance(file_obj, io.BytesIO):
            # BytesIO from JSON mode
            file_obj.seek(0)
            return file_obj.read()
        else:
            # FileStorage from multipart mode
            file_obj.seek(0)
            return file_obj.read()
