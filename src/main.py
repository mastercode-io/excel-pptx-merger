"""Main API handler with Flask and Google Cloud Function support - Enhanced with image position support."""

import json
import logging
import os
import traceback
import datetime
import copy
import base64
import io
from typing import Any, Dict, List, Optional, Tuple, Union
from flask import Flask, request, jsonify, send_file
from werkzeug.exceptions import RequestEntityTooLarge
import click
import functions_framework
import shutil
import uuid

from .config_manager import ConfigManager
from .excel_processor import ExcelProcessor
from .pptx_processor import PowerPointProcessor
from .excel_updater import ExcelUpdater
from .temp_file_manager import TempFileManager
from .utils.exceptions import (
    ExcelPptxMergerError,
    ValidationError,
    FileProcessingError,
    ConfigurationError,
    APIError,
    AuthenticationError,
    ExcelProcessingError,
)
from .utils.validation import validate_api_request
from .utils.file_utils import save_uploaded_file, get_file_info

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize components
config_manager = ConfigManager()
app_config = config_manager.get_app_config()

# Configure Flask app
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = app_config["max_file_size_mb"] * 1024 * 1024
app.config["MAX_FORM_MEMORY_SIZE"] = 10 * 1024 * 1024  # 10MB for form fields (was 500KB default)


def setup_logging() -> None:
    """Setup logging configuration."""
    log_level = app_config.get("log_level", "INFO").upper()
    logging.getLogger().setLevel(getattr(logging, log_level, logging.INFO))

    # Set up format
    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )

    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logging.getLogger().addHandler(console_handler)
    
    # Log Flask configuration limits
    max_content_length = app.config.get("MAX_CONTENT_LENGTH", 0)
    max_form_memory_size = app.config.get("MAX_FORM_MEMORY_SIZE", 0)
    
    logger.info(f"Flask configuration - MAX_CONTENT_LENGTH: {max_content_length / (1024*1024):.1f}MB")
    logger.info(f"Flask configuration - MAX_FORM_MEMORY_SIZE: {max_form_memory_size / (1024*1024):.1f}MB")


def authenticate_request() -> bool:
    """Authenticate API request."""
    # Skip authentication in development mode
    if app_config.get("development_mode", False):
        logger.debug("Authentication bypassed in development mode")
        return True

    api_key = app_config.get("api_key")
    if not api_key:
        return True  # No authentication required if no key configured

    # Check Authorization header
    auth_header = request.headers.get("Authorization")
    if auth_header and auth_header.startswith("Bearer "):
        token = auth_header.split(" ", 1)[1]
        return token == api_key

    # Check query parameter
    request_key = request.args.get("api_key") or request.form.get("api_key")
    return request_key == api_key


def create_error_response(
    error: Exception, status_code: int = 500
) -> Tuple[Dict[str, Any], int]:
    """Create standardized error response."""
    error_response = {
        "success": False,
        "error": {
            "type": type(error).__name__,
            "message": str(error),
            "code": status_code,
        },
    }

    # Add error code if available
    if hasattr(error, "error_code") and error.error_code:
        error_response["error"]["error_code"] = error.error_code

    # Add traceback in development mode
    if app_config.get("development_mode", False):
        error_response["error"]["traceback"] = traceback.format_exc()

    logger.error(f"API Error ({status_code}): {error}")
    return error_response, status_code


@app.errorhandler(RequestEntityTooLarge)
def handle_file_too_large(error):
    """Handle file size too large error with detailed diagnostics."""
    # Enhanced logging for 413 errors
    content_type = request.headers.get('Content-Type', 'Unknown')
    content_length = request.headers.get('Content-Length', 'Unknown')
    user_agent = request.headers.get('User-Agent', 'Unknown')
    request_path = request.path
    
    logger.error(f"413 Error - Request too large:")
    logger.error(f"  Path: {request_path}")
    logger.error(f"  Content-Type: {content_type}")
    logger.error(f"  Content-Length header: {content_length}")
    logger.error(f"  User-Agent: {user_agent}")
    logger.error(f"  Max allowed size: {app_config['max_file_size_mb']}MB")
    
    # Calculate size in MB if content-length is available
    size_info = ""
    if content_length != 'Unknown' and content_length.isdigit():
        size_mb = int(content_length) / (1024 * 1024)
        size_info = f" (Request size: {size_mb:.2f}MB)"
        logger.error(f"  Actual request size: {size_mb:.2f}MB")
    
    # Check if this might be a CRM-specific issue
    is_likely_crm = False
    if user_agent != 'Unknown':
        crm_indicators = ['deluge', 'zoho', 'crm', 'automation', 'webhook']
        is_likely_crm = any(indicator in user_agent.lower() for indicator in crm_indicators)
    
    if is_likely_crm:
        logger.error("  Potential CRM/automation request detected")
    
    # Create enhanced error message
    error_message = f"Request size exceeds maximum allowed size of {app_config['max_file_size_mb']}MB{size_info}"
    
    if is_likely_crm:
        error_message += ". Note: CRM systems may inflate request size due to base64 encoding or additional metadata."
    
    return create_error_response(
        ValidationError(error_message),
        413,
    )


@app.before_request
def before_request():
    """Pre-request authentication and validation."""
    # Skip authentication for health check
    if request.endpoint == "health":
        return

    # Authenticate request
    if not authenticate_request():
        error_response, status_code = create_error_response(
            AuthenticationError("Invalid API key"), 401
        )
        return jsonify(error_response), status_code


@app.route("/api/v1/health", methods=["GET"])
def health() -> Tuple[Dict[str, Any], int]:
    """Health check endpoint."""
    try:
        health_info = {
            "success": True,
            "status": "healthy",
            "version": "0.1.0",
            "environment": os.getenv("ENVIRONMENT", "development"),
            "timestamp": None,
            "services": {"config_manager": True, "temp_file_manager": True},
            "features": {
                "enhanced_image_extraction": True,
                "position_based_image_matching": True,
                "debug_mode": app_config.get("development_mode", False),
            },
        }

        # Add timestamp
        from datetime import datetime

        health_info["timestamp"] = datetime.utcnow().isoformat() + "Z"

        return health_info, 200

    except Exception as e:
        return create_error_response(e, 500)


@app.route("/api/v1/merge", methods=["POST"])
def merge_files() -> Union[Tuple[Dict[str, Any], int], Any]:
    """Main file processing endpoint with enhanced image support."""
    try:
        # Get files from request
        excel_file = request.files.get("excel_file")
        pptx_file = request.files.get("pptx_file")

        # Check if files are provided
        if not excel_file or not pptx_file:
            return create_error_response(
                ValidationError("Excel and PowerPoint files are required"), 400
            )

        # Get session ID from headers or generate a new one (needed early for auto-detection)
        session_id = request.headers.get("X-Session-ID")
        if not session_id:
            session_id = str(uuid.uuid4())
            logger.info(f"Generated new session ID: {session_id}")
        else:
            logger.info(f"Using provided session ID: {session_id}")

        # Check if we should save files
        save_files = app_config.get("save_files", False)
        logger.info(
            f"File saving mode: {'enabled' if save_files else 'disabled (memory-only)'}"
        )

        # Get extraction configuration
        extraction_config = {}
        if "config" in request.form:
            try:
                extraction_config = json.loads(request.form.get("config", "{}"))
            except json.JSONDecodeError as e:
                logger.error(f"Invalid JSON configuration: {e}")
                return jsonify({"error": f"Invalid JSON configuration: {e}"}), 400
        elif request.is_json:
            extraction_config = request.json or {}

        # If no configuration provided, use auto-detection
        if not extraction_config:
            logger.info("No configuration provided, using auto-detection")
            excel_processor_for_detection = ExcelProcessor(excel_file)

            try:
                extraction_config = (
                    excel_processor_for_detection.auto_detect_all_sheets()
                )
                logger.info("Using auto-detection for all sheets in merge operation")
            except Exception as e:
                logger.error(f"Auto-detection failed: {e}")
                return create_error_response(
                    ExcelProcessingError(f"Failed to auto-detect Excel structure: {e}"),
                    500,
                )
            finally:
                excel_processor_for_detection.close()

        logger.debug(f"Extraction config: {extraction_config}")

        # Initialize file handling
        temp_manager = None
        temp_dir = None
        excel_path = None
        pptx_path = None

        if save_files:
            # Traditional file-based processing
            temp_manager = TempFileManager()
            temp_dir = temp_manager.get_session_directory(session_id)
            logger.info(f"Using session directory: {temp_dir}")

            # Save uploaded files to temp directory
            excel_path = temp_manager.save_file_to_temp(
                temp_dir, excel_file.filename, excel_file, temp_manager.FILE_TYPE_INPUT
            )
            pptx_path = temp_manager.save_file_to_temp(
                temp_dir, pptx_file.filename, pptx_file, temp_manager.FILE_TYPE_INPUT
            )

            logger.info(f"Saved input files to: {excel_path}, {pptx_path}")
        else:
            # Memory-only processing
            logger.info("Processing files in memory without saving to disk")

        # Process Excel file
        if save_files:
            excel_processor = ExcelProcessor(excel_path)
        else:
            excel_processor = ExcelProcessor(excel_file)

        try:
            try:
                # Extract data from Excel
                extracted_data = excel_processor.extract_data(
                    extraction_config.get("global_settings", {}),
                    extraction_config.get("sheet_configs", {}),
                )
                logger.info(f"Successfully extracted data from Excel file")
            except Exception as e:
                logger.error(f"Failed to extract data from Excel: {e}")
                return (
                    jsonify({"error": f"Failed to extract data from Excel: {e}"}),
                    500,
                )

            # Extract images from Excel file
            logger.info("Extracting images from Excel file")
            if save_files:
                images = excel_processor.extract_images(temp_dir)
            else:
                images = excel_processor.extract_images()

            # Log the number of images extracted
            image_count = sum(len(sheet_images) for sheet_images in images.values())
            logger.info(f"Extracted {image_count} images from Excel file")

            # Verify images were extracted
            if image_count == 0:
                logger.warning("No images were extracted from the Excel file")
            else:
                # Link images to the extracted data
                if (
                    extraction_config.get("global_settings", {})
                    .get("image_extraction", {})
                    .get("enabled", True)
                ):
                    extracted_data = excel_processor.link_images_to_table(
                        extracted_data, images
                    )
                    logger.info("Linked images to extracted data")

                # Log the image paths for debugging
                for sheet_name, sheet_images in images.items():
                    logger.info(f"Sheet {sheet_name} has {len(sheet_images)} images")
                    for img in sheet_images:
                        logger.debug(f"Image path: {img['path']}")
                        # Verify the image file exists
                        if os.path.exists(img["path"]):
                            logger.debug(f"Verified image exists at: {img['path']}")
                        else:
                            logger.error(f"Image file does not exist at: {img['path']}")
        finally:
            excel_processor.close()

        # Process PowerPoint file
        if save_files:
            pptx_processor = PowerPointProcessor(pptx_path)
        else:
            pptx_processor = PowerPointProcessor(pptx_file)

        try:
            output_filename = f"merged_{os.path.basename(pptx_file.filename)}"

            if save_files:
                # File-based processing: save to disk
                merged_file_path = temp_manager.storage.get_output_path(
                    temp_dir, output_filename
                )

                # Ensure output directory exists
                os.makedirs(os.path.dirname(merged_file_path), exist_ok=True)

                # Merge data into PowerPoint and save
                merged_file_path = pptx_processor.merge_data(
                    extracted_data, merged_file_path, images
                )
            else:
                # Memory-based processing: create in-memory file
                import tempfile

                with tempfile.NamedTemporaryFile(
                    suffix=".pptx", delete=False
                ) as tmp_file:
                    merged_file_path = tmp_file.name

                # Merge data into PowerPoint and save to temporary file
                merged_file_path = pptx_processor.merge_data(
                    extracted_data, merged_file_path, images
                )

            # Verify the merged file exists and ensure it's an absolute path
            if not os.path.isabs(merged_file_path):
                merged_file_path = os.path.abspath(merged_file_path)

            if not os.path.exists(merged_file_path):
                logger.error(f"Merged file does not exist at path: {merged_file_path}")
                raise FileNotFoundError(f"Merged file not found at: {merged_file_path}")
            else:
                logger.debug(f"Verified merged file exists at: {merged_file_path}")

            # In development mode, also save a copy to the debug folder (only if saving files)
            if app_config.get("development_mode", False) and save_files and temp_dir:
                # Ensure debug directory exists with absolute path
                debug_dir = os.path.join(temp_dir, temp_manager.FILE_TYPE_DEBUG)
                if not os.path.isabs(debug_dir):
                    debug_dir = os.path.abspath(debug_dir)
                os.makedirs(debug_dir, exist_ok=True)
                logger.info(f"Ensuring debug directory exists: {debug_dir}")

                # Save the extracted data to a JSON file for debugging
                debug_data_filename = f"debug_data_{session_id}.json"
                debug_data_path = os.path.join(debug_dir, debug_data_filename)
                try:
                    with open(debug_data_path, "w") as f:
                        json.dump(extracted_data, f, indent=2, default=str)
                    logger.info(f"Saved extracted data to: {debug_data_path}")
                except Exception as e:
                    logger.error(f"Failed to save debug data: {e}")

                # Use direct file copy instead of reading/writing through temp_manager
                debug_output_filename = f"debug_{output_filename}"
                debug_output_path = os.path.join(debug_dir, debug_output_filename)

                # Copy the file directly
                try:
                    # Verify source file exists
                    if os.path.exists(merged_file_path):
                        shutil.copy2(merged_file_path, debug_output_path)
                        logger.info(
                            f"Development mode: Saved copy of merged file to {debug_output_path}"
                        )

                        # Verify the debug file was created
                        if os.path.exists(debug_output_path):
                            logger.debug(
                                f"Debug file successfully created at: {debug_output_path}"
                            )
                        else:
                            logger.error(
                                f"Failed to create debug file at: {debug_output_path}"
                            )
                    else:
                        logger.error(
                            f"Cannot copy to debug: Source file does not exist at {merged_file_path}"
                        )
                except Exception as e:
                    logger.error(f"Failed to save debug copy: {e}")
            elif app_config.get("development_mode", False) and not save_files:
                logger.info(
                    "Development mode: Debug file saving skipped (memory-only mode)"
                )

            # Clean up images after successful merge if configured
            if images:
                excel_processor.cleanup_images(
                    images, extraction_config.get("global_settings", {})
                )

            # Return the merged file
            # Use the absolute path directly to avoid path resolution issues
            logger.debug(f"Sending file with absolute path: {merged_file_path}")

            def cleanup_temp_file():
                """Cleanup temporary file after response is sent (for memory-only mode)."""
                if (
                    not save_files
                    and merged_file_path
                    and os.path.exists(merged_file_path)
                ):
                    try:
                        os.unlink(merged_file_path)
                        logger.debug(f"Cleaned up temporary file: {merged_file_path}")
                    except Exception as e:
                        logger.warning(f"Failed to cleanup temporary file: {e}")

            response = send_file(
                path_or_file=merged_file_path,  # Use the verified absolute path
                as_attachment=True,
                download_name=output_filename,
                mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )

            # Schedule cleanup for memory-only mode
            if not save_files:
                import threading

                cleanup_thread = threading.Thread(target=cleanup_temp_file)
                cleanup_thread.daemon = True
                cleanup_thread.start()

            return response
        finally:
            pptx_processor.close()

    except Exception as e:
        logger.exception(f"Error in merge endpoint: {e}")
        return jsonify({"error": str(e)}), 500


@app.route("/api/v1/preview", methods=["POST"])
def preview_merge() -> Tuple[Dict[str, Any], int]:
    """Preview merge without performing actual merge - Enhanced with image analysis."""
    temp_manager = None
    temp_dir = None

    try:
        # Validate request
        if "excel_file" not in request.files or "pptx_file" not in request.files:
            raise ValidationError("Both 'excel_file' and 'pptx_file' are required")

        excel_file = request.files["excel_file"]
        pptx_file = request.files["pptx_file"]

        # Get configuration
        config_data = request.form.get("config")
        if config_data:
            extraction_config = json.loads(config_data)
        else:
            extraction_config = config_manager.get_default_config()

        # Initialize temp file manager
        temp_file_config = extraction_config.get("global_settings", {}).get(
            "temp_file_cleanup", {}
        )
        temp_manager = TempFileManager(temp_file_config)
        temp_dir = temp_manager.create_temp_directory()

        # Save files
        excel_path = temp_manager.save_file_to_temp(
            temp_dir, excel_file.filename, excel_file
        )
        pptx_path = temp_manager.save_file_to_temp(
            temp_dir, pptx_file.filename, pptx_file
        )

        # Process Excel file
        excel_processor = ExcelProcessor(excel_path)
        try:
            extracted_data = excel_processor.extract_data(
                extraction_config.get("global_settings", {}),
                extraction_config.get("sheet_configs", {}),
            )

            # Extract images with position information
            images = None
            if (
                extraction_config.get("global_settings", {})
                .get("image_extraction", {})
                .get("enabled", True)
            ):
                images = excel_processor.extract_images(temp_dir)
        finally:
            excel_processor.close()

        # Analyze PowerPoint template
        pptx_processor = PowerPointProcessor(pptx_path)
        try:
            template_info = pptx_processor.validate_template()
            preview_info = pptx_processor.preview_merge(extracted_data, images)
            image_requirements = pptx_processor.get_image_requirements()
        finally:
            pptx_processor.close()

        # Compile enhanced preview response
        preview_response = {
            "success": True,
            "preview": {
                "extracted_data": extracted_data,
                "template_info": template_info,
                "merge_preview": preview_info,
                "image_analysis": {
                    "extracted_images": _create_image_summary(images) if images else {},
                    "image_requirements": image_requirements,
                    "matching_analysis": _analyze_image_matching(
                        images, preview_info.get("image_placeholders", [])
                    ),
                },
                "configuration_used": extraction_config,
            },
        }

        return preview_response, 200

    except Exception as e:
        if temp_manager and temp_dir:
            temp_manager.mark_error(temp_dir)
        return create_error_response(e, 500)

    finally:
        if temp_manager and temp_dir:
            temp_manager.schedule_cleanup(temp_dir)


def _create_image_summary(images: Dict[str, List[Dict[str, Any]]]) -> Dict[str, Any]:
    """Create summary of extracted images for preview."""
    if not images:
        return {}

    summary = {
        "total_images": 0,
        "sheets": {},
        "position_info": {
            "images_with_position": 0,
            "images_without_position": 0,
            "anchor_types": {},
        },
    }

    for sheet_name, sheet_images in images.items():
        sheet_summary = {"count": len(sheet_images), "images": []}

        for image_info in sheet_images:
            image_summary = {
                "filename": image_info.get("filename"),
                "index": image_info.get("index"),
                "format": image_info.get("format"),
                "size": image_info.get("size"),
                "position": image_info.get("position", {}),
            }

            # Count position statistics
            position = image_info.get("position", {})
            if position.get("estimated_cell"):
                summary["position_info"]["images_with_position"] += 1
                anchor_type = position.get("anchor_type", "unknown")
                summary["position_info"]["anchor_types"][anchor_type] = (
                    summary["position_info"]["anchor_types"].get(anchor_type, 0) + 1
                )
            else:
                summary["position_info"]["images_without_position"] += 1

            sheet_summary["images"].append(image_summary)

        summary["sheets"][sheet_name] = sheet_summary
        summary["total_images"] += len(sheet_images)

    return summary


def _analyze_image_matching(
    images: Optional[Dict[str, List[Dict[str, Any]]]], placeholders: List[str]
) -> Dict[str, Any]:
    """Analyze potential image-placeholder matching."""
    analysis = {
        "total_placeholders": len(placeholders),
        "potential_matches": {},
        "unmatched_placeholders": [],
        "unmatched_images": [],
        "recommendations": [],
    }

    if not images or not placeholders:
        analysis["unmatched_placeholders"] = placeholders.copy()
        if images:
            all_images = []
            for sheet_images in images.values():
                all_images.extend(sheet_images)
            analysis["unmatched_images"] = [img.get("filename") for img in all_images]
        return analysis

    # Create flat list of all images with their info
    all_images = []
    for sheet_name, sheet_images in images.items():
        for image_info in sheet_images:
            all_images.append({"sheet": sheet_name, "info": image_info})

    matched_images = set()

    # Try to match placeholders with images
    for placeholder in placeholders:
        best_match = None
        match_confidence = 0

        for idx, image_entry in enumerate(all_images):
            if idx in matched_images:
                continue

            confidence = _calculate_match_confidence(placeholder, image_entry["info"])
            if confidence > match_confidence:
                match_confidence = confidence
                best_match = {
                    "image_index": idx,
                    "confidence": confidence,
                    "image_info": image_entry,
                }

        if best_match and match_confidence > 0.3:  # Threshold for reasonable match
            analysis["potential_matches"][placeholder] = best_match
            matched_images.add(best_match["image_index"])
        else:
            analysis["unmatched_placeholders"].append(placeholder)

    # Find unmatched images
    for idx, image_entry in enumerate(all_images):
        if idx not in matched_images:
            analysis["unmatched_images"].append(image_entry["info"].get("filename"))

    # Generate recommendations
    analysis["recommendations"] = _generate_matching_recommendations(analysis)

    return analysis


def _calculate_match_confidence(placeholder: str, image_info: Dict[str, Any]) -> float:
    """Calculate confidence score for placeholder-image matching."""
    confidence = 0.0
    placeholder_lower = placeholder.lower()

    # Position-based matching (highest confidence)
    position = image_info.get("position", {})
    if position.get("estimated_cell"):
        cell_ref = position["estimated_cell"].lower()
        if cell_ref in placeholder_lower or placeholder_lower.endswith(cell_ref):
            confidence += 0.8

    # Index-based matching
    import re

    placeholder_numbers = re.findall(r"\d+", placeholder_lower)
    image_index = image_info.get("index", 0)

    if placeholder_numbers:
        for num_str in placeholder_numbers:
            try:
                num = int(num_str)
                if (
                    num == image_index or num == image_index - 1
                ):  # 0-based or 1-based indexing
                    confidence += 0.6
                    break
            except ValueError:
                continue

    # Keyword matching
    keywords = ["image", "img", "picture", "photo"]
    for keyword in keywords:
        if keyword in placeholder_lower:
            confidence += 0.3
            break

    # Sheet name matching
    sheet_name = image_info.get("sheet", "").lower()
    if sheet_name and sheet_name.replace(" ", "_") in placeholder_lower:
        confidence += 0.4

    return min(confidence, 1.0)  # Cap at 1.0


def _generate_matching_recommendations(analysis: Dict[str, Any]) -> List[str]:
    """Generate recommendations for improving image matching."""
    recommendations = []

    unmatched_count = len(analysis["unmatched_placeholders"])
    unmatched_images_count = len(analysis["unmatched_images"])

    if unmatched_count > 0:
        recommendations.append(
            f"Consider updating {unmatched_count} placeholder(s) to include position information "
            f"(e.g., {{{{image_A1}}}} for image at cell A1)"
        )

    if unmatched_images_count > 0:
        recommendations.append(
            f"{unmatched_images_count} extracted image(s) could not be matched to placeholders. "
            f"Consider adding corresponding placeholders in your PowerPoint template."
        )

    if analysis["potential_matches"]:
        low_confidence_matches = [
            placeholder
            for placeholder, match in analysis["potential_matches"].items()
            if match["confidence"] < 0.6
        ]

        if low_confidence_matches:
            recommendations.append(
                f"Consider improving placeholder names for better matching: {', '.join(low_confidence_matches)}"
            )

    if not recommendations:
        recommendations.append(
            "Image matching analysis looks good! All placeholders have potential matches."
        )

    return recommendations


@app.route("/api/v1/config", methods=["GET", "POST"])
def manage_config() -> Tuple[Dict[str, Any], int]:
    """Manage configuration with enhanced image extraction settings."""
    try:
        if request.method == "GET":
            # Return default configuration with enhanced image settings
            config = config_manager.get_default_config()
            return {
                "success": True,
                "config": config,
                "features": {
                    "enhanced_image_extraction": True,
                    "position_based_matching": True,
                    "supported_image_formats": ["png", "jpg", "jpeg", "gif", "webp"],
                },
            }, 200

        elif request.method == "POST":
            # Validate and save configuration
            config_data = request.get_json()
            if not config_data:
                raise ValidationError("Configuration data is required")

            config_manager.validate_runtime_config(config_data)

            # For now, just validate - future versions could store custom configs
            return {
                "success": True,
                "message": "Configuration validated successfully",
                "config": config_data,
            }, 200

    except Exception as e:
        return create_error_response(e, 400)


@app.route("/api/v1/extract", methods=["POST"])
def extract_data_endpoint() -> Union[Tuple[Dict[str, Any], int], Any]:
    """Extract data from specified Excel sheets and return as JSON."""
    start_time = datetime.datetime.now()

    try:
        # Validate request
        if "excel_file" not in request.files:
            return create_error_response(ValidationError("Excel file is required"), 400)

        excel_file = request.files["excel_file"]

        # Get sheet_names parameter (required)
        sheet_names_param = request.form.get("sheet_names")
        if not sheet_names_param:
            return create_error_response(
                ValidationError("sheet_names parameter is required"), 400
            )

        # Parse sheet_names - should be a JSON array of strings
        try:
            sheet_names = json.loads(sheet_names_param)
            if not isinstance(sheet_names, list):
                return create_error_response(
                    ValidationError("sheet_names must be a list of strings"), 400
                )

            # Validate all items are strings
            for name in sheet_names:
                if not isinstance(name, str):
                    return create_error_response(
                        ValidationError("All items in sheet_names must be strings"), 400
                    )

            if not sheet_names:
                return create_error_response(
                    ValidationError("sheet_names list cannot be empty"), 400
                )

        except json.JSONDecodeError as e:
            return create_error_response(
                ValidationError(f"Invalid JSON in sheet_names parameter: {e}"), 400
            )

        # Get configuration (optional)
        config = None
        if "config" in request.form:
            try:
                config = json.loads(request.form.get("config"))
            except json.JSONDecodeError as e:
                return create_error_response(
                    ValidationError(f"Invalid JSON configuration: {e}"), 400
                )

        # Get auto-detect setting
        auto_detect = request.form.get("auto_detect", "true").lower() == "true"

        # Get max_rows parameter (optional)
        max_rows = None
        if "max_rows" in request.form:
            try:
                max_rows = int(request.form.get("max_rows"))
                if max_rows <= 0:
                    return create_error_response(
                        ValidationError("max_rows must be a positive integer"), 400
                    )
            except ValueError:
                return create_error_response(
                    ValidationError("max_rows must be a valid integer"), 400
                )

        logger.info(
            f"Extracting data from sheets {sheet_names} with auto_detect={auto_detect}, max_rows={max_rows}"
        )

        # Process Excel file (use existing memory/file handling logic)
        excel_processor = ExcelProcessor(excel_file)
        try:
            # Get available sheet names for validation
            available_sheets = excel_processor.get_sheet_names()

            # Process each sheet
            tabs = []
            for sheet_name in sheet_names:
                # Check if sheet exists
                if sheet_name not in available_sheets:
                    logger.warning(
                        f"Sheet '{sheet_name}' not found in workbook. Available sheets: {available_sheets}"
                    )
                    # Add empty result for missing sheet
                    tab_result = {
                        "success": True,
                        "sheet_name": sheet_name,
                        "extracted_data": {},
                        "metadata": {
                            "total_rows": 0,
                            "extracted_rows": 0,
                            "total_columns": 0,
                            "extraction_method": "sheet_not_found",
                            "data_types_detected": [],
                            "error": f"Sheet '{sheet_name}' not found",
                        },
                    }
                    tabs.append(tab_result)
                    continue

                # Get per-sheet configuration if provided
                sheet_config = None
                if config and isinstance(config, dict):
                    # Look for sheet-specific config
                    sheet_config = config.get(sheet_name) or config.get("default")

                # Log what configuration we're using
                logger.debug(
                    f"Sheet {sheet_name}: config={sheet_config}, auto_detect={auto_detect}"
                )

                # If no configuration provided, let auto-detection handle it
                # (auto_detect=True is already set by default above)

                try:
                    # Extract data from this sheet
                    extracted_data = excel_processor.extract_single_sheet(
                        sheet_name=sheet_name,
                        config=sheet_config,
                        auto_detect=auto_detect,
                        max_rows=max_rows,
                    )

                    # Build tab result in the same format as before
                    tab_result = {
                        "success": True,
                        "sheet_name": sheet_name,
                        "extracted_data": extracted_data["data"],
                        "metadata": {
                            "total_rows": extracted_data["metadata"]["total_rows"],
                            "extracted_rows": extracted_data["metadata"][
                                "extracted_rows"
                            ],
                            "total_columns": extracted_data["metadata"][
                                "total_columns"
                            ],
                            "extraction_method": extracted_data["metadata"]["method"],
                            "data_types_detected": extracted_data["metadata"]["types"],
                        },
                    }

                except Exception as e:
                    logger.error(
                        f"Error extracting data from sheet '{sheet_name}': {e}"
                    )
                    # Add error result for this sheet
                    tab_result = {
                        "success": False,
                        "sheet_name": sheet_name,
                        "extracted_data": {},
                        "metadata": {
                            "total_rows": 0,
                            "extracted_rows": 0,
                            "total_columns": 0,
                            "extraction_method": "error",
                            "data_types_detected": [],
                            "error": str(e),
                        },
                    }

                tabs.append(tab_result)

            # Calculate processing time
            processing_time = (
                datetime.datetime.now() - start_time
            ).total_seconds() * 1000

            # Build response with tabs structure
            response = {
                "success": True,
                "tabs": tabs,
                "summary": {
                    "total_sheets_requested": len(sheet_names),
                    "sheets_processed": len(
                        [
                            tab
                            for tab in tabs
                            if tab["success"]
                            and tab["metadata"].get("extraction_method")
                            != "sheet_not_found"
                        ]
                    ),
                    "sheets_not_found": len(
                        [
                            tab
                            for tab in tabs
                            if tab["metadata"].get("extraction_method")
                            == "sheet_not_found"
                        ]
                    ),
                    "sheets_with_errors": len(
                        [tab for tab in tabs if not tab["success"]]
                    ),
                    "timestamp": datetime.datetime.utcnow().isoformat() + "Z",
                    "processing_time_ms": round(processing_time, 2),
                },
            }

            return response, 200

        finally:
            excel_processor.close()

    except ValidationError as e:
        logger.warning(f"Validation error in extract endpoint: {e}")
        return create_error_response(e, 400)
    except Exception as e:
        logger.exception(f"Error in extract endpoint: {e}")
        return create_error_response(e, 500)


@app.route("/api/v1/stats", methods=["GET"])
def get_stats() -> Tuple[Dict[str, Any], int]:
    """Get system statistics with enhanced feature information."""
    try:
        # This could be expanded to include more detailed stats
        stats = {
            "success": True,
            "stats": {
                "app_config": {
                    "max_file_size_mb": app_config["max_file_size_mb"],
                    "allowed_extensions": app_config["allowed_extensions"],
                    "development_mode": app_config["development_mode"],
                },
                "features": {
                    "enhanced_image_extraction": True,
                    "position_based_image_matching": True,
                    "image_format_detection": True,
                    "debug_mode": app_config.get("development_mode", False),
                },
                "runtime": {
                    "python_version": os.sys.version,
                    "environment": os.getenv("ENVIRONMENT", "development"),
                },
            },
        }

        return stats, 200

    except Exception as e:
        return create_error_response(e, 500)


@app.route("/api/v1/update", methods=["POST"])
def update_excel_file() -> Union[Tuple[Dict[str, Any], int], Any]:
    """Update Excel file with provided data - supports both multipart and JSON modes."""
    temp_manager = None
    
    try:
        # Enhanced request logging for debugging
        content_type = request.headers.get('Content-Type', 'Not specified')
        content_length = request.headers.get('Content-Length', 'Not specified')
        user_agent = request.headers.get('User-Agent', 'Not specified')
        
        logger.info(f"Update request received - Content-Type: {content_type}, Content-Length: {content_length}")
        logger.info(f"Update request User-Agent: {user_agent}")
        
        # Enhanced dual mode detection - handle incorrect Content-Type from CRM systems
        is_json_request = request.is_json
        has_form_data = bool(request.form)
        has_files = bool(request.files)
        
        # Fallback JSON detection for CRM systems that send JSON with wrong Content-Type
        if not is_json_request and request.data and not has_form_data and not has_files:
            try:
                # Try to parse raw request data as JSON
                json.loads(request.data)
                is_json_request = True
                logger.info(f"Detected JSON payload despite Content-Type: {content_type}")
            except json.JSONDecodeError:
                logger.debug("Raw request data is not valid JSON")
        
        logger.info(f"Request analysis - JSON: {is_json_request}, Form data: {has_form_data}, Files: {has_files}")
        
        # Log Content-Type detection issues
        if is_json_request and not request.is_json:
            logger.warning(f"JSON payload detected with non-standard Content-Type: {content_type}")
        elif content_type.startswith('text/plain') and request.data:
            logger.info(f"text/plain Content-Type with {len(request.data)} bytes of data")
        
        # Log request size breakdown
        if hasattr(request, 'content_length') and request.content_length:
            logger.info(f"Actual request content length: {request.content_length} bytes ({request.content_length / (1024*1024):.2f} MB)")
        
        # Initialize variables for unified processing
        excel_file = None
        excel_data = None
        update_data = None
        config = None
        include_update_log = False
        
        if is_json_request:
            # JSON MODE: Everything as base64 strings
            logger.info("Processing request in JSON mode (base64 Excel file)")
            
            try:
                # Handle both standard JSON requests and CRM systems with wrong Content-Type
                if request.is_json:
                    json_data = request.get_json()
                else:
                    # Parse raw data for systems that send JSON with text/plain Content-Type
                    json_data = json.loads(request.data)
                    logger.info("Parsed JSON from raw request data due to incorrect Content-Type")
                
                if not json_data:
                    return create_error_response(
                        ValidationError("JSON payload is required"), 400
                    )
                
                # Extract Excel file from base64
                excel_file_b64 = json_data.get('excel_file')
                if not excel_file_b64:
                    return create_error_response(
                        ValidationError("excel_file (base64) is required in JSON mode"), 400
                    )
                
                logger.info(f"Base64 Excel file size: {len(excel_file_b64)} characters")
                
                # Decode base64 Excel file
                try:
                    excel_data = base64.b64decode(excel_file_b64)
                    excel_file = io.BytesIO(excel_data)
                    excel_file.filename = json_data.get('filename', 'uploaded_file.xlsx')
                    logger.info(f"Decoded Excel file size: {len(excel_data)} bytes")
                except Exception as e:
                    return create_error_response(
                        ValidationError(f"Invalid base64 Excel file: {e}"), 400
                    )
                
                # Extract update data, config, and include_update_log directly from JSON
                update_data = json_data.get('update_data', {})
                config = json_data.get('config', {})
                include_update_log = json_data.get('include_update_log', False)
                
                logger.info(f"JSON mode - Update data fields: {list(update_data.keys()) if isinstance(update_data, dict) else 'not a dict'}")
                logger.info(f"JSON mode - Include update log: {include_update_log}")
                logger.info("JSON mode processing completed successfully")
                
            except json.JSONDecodeError as e:
                return create_error_response(
                    ValidationError(f"Invalid JSON format: {e}"), 400
                )
                
        else:
            # MULTIPART MODE: Binary file + JSON form fields
            logger.info("Processing request in multipart mode (binary Excel file)")
            
            # Log form fields for debugging
            if has_form_data:
                form_fields = list(request.form.keys())
                logger.info(f"Form fields present: {form_fields}")
                for field in form_fields:
                    field_size = len(request.form.get(field, ''))
                    logger.info(f"Form field '{field}' size: {field_size} bytes")
            
            # Log file information
            if has_files:
                file_info = []
                for file_key in request.files.keys():
                    file_obj = request.files[file_key]
                    file_info.append(f"{file_key}: {file_obj.filename}")
                logger.info(f"Files in request: {file_info}")
            
            # Validate Excel file presence
            if "excel_file" not in request.files:
                logger.error("Excel file missing from multipart request")
                return create_error_response(
                    ValidationError("Excel file is required"), 400
                )
                
            excel_file = request.files["excel_file"]
            
            # Get update data, config, and include_update_log from form fields
            update_data_str = request.form.get("update_data", "{}")
            config_str = request.form.get("config", "{}")
            include_update_log_str = request.form.get("include_update_log", "false")
            
            logger.info(f"Update data string size: {len(update_data_str)} bytes")
            logger.info(f"Config string size: {len(config_str)} bytes")
            logger.info(f"Include update log: {include_update_log_str}")
            
            # Parse JSON from form fields
            try:
                update_data = json.loads(update_data_str)
                config = json.loads(config_str)
                include_update_log = include_update_log_str.lower() in ('true', '1', 'yes', 'on')
                logger.info("Successfully parsed JSON data from form fields")
            except json.JSONDecodeError as e:
                logger.error(f"JSON parsing failed: {e}")
                logger.error(f"Update data preview: {update_data_str[:200]}...")
                logger.error(f"Config preview: {config_str[:200]}...")
                return create_error_response(
                    ValidationError(f"Invalid JSON format: {e}"), 400
                )
        
        # UNIFIED PROCESSING PATH: Both modes now have same data structure
        # At this point we have:
        # - excel_file: file-like object (either uploaded file or BytesIO from base64)
        # - update_data: dict with data to update
        # - config: dict with configuration
        
        # Validate file
        if not excel_file or not hasattr(excel_file, 'filename') or not excel_file.filename:
            return create_error_response(
                ValidationError("No file selected or invalid file"), 400
            )
            
        logger.info(f"Processing Excel update request for file: {excel_file.filename}")
        logger.info(f"Mode: {'JSON (base64)' if is_json_request else 'Multipart (binary)'}")
        
        # Setup temp file management
        temp_manager = TempFileManager(app_config["temp_file_cleanup"])
        temp_dir = temp_manager.create_temp_directory()
        
        # Save Excel file to temp directory (works for both modes)
        excel_filename = f"input_{uuid.uuid4().hex[:8]}.xlsx"
        excel_path = os.path.join(temp_dir, excel_filename)
        
        if is_json_request:
            # For JSON mode: write decoded bytes to file
            with open(excel_path, 'wb') as f:
                f.write(excel_data)
            logger.info(f"Saved base64-decoded Excel file to: {excel_path}")
        else:
            # For multipart mode: save uploaded file
            excel_file.save(excel_path)
            logger.info(f"Saved uploaded Excel file to: {excel_path}")
        
        # Update Excel file (same process for both modes)
        updater = ExcelUpdater(excel_path)
        try:
            updated_path = updater.update_excel(update_data, config, include_update_log)
            
            # Return updated file
            return send_file(
                updated_path,
                as_attachment=True,
                download_name=f"updated_{excel_file.filename}",
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            logger.error(f"Excel update failed: {e}")
            return create_error_response(e, 400)
        finally:
            updater.close()
        
    except Exception as e:
        logger.error(f"Update endpoint error: {e}")
        return create_error_response(e, 500)
    finally:
        if temp_manager:
            temp_manager.cleanup_old_directories()


def save_debug_info(extracted_data, images, temp_dir, base_filename):
    """Save enhanced debug information with image positions to files for development purposes."""
    if not app_config.get("development_mode", False):
        return None

    # Create a deep copy of the extracted data to ensure we don't modify the original
    debug_data = copy.deepcopy(extracted_data)

    # Process image structure for debug data
    if images:
        image_refs = {}
        for sheet_name, sheet_images in images.items():
            sheet_image_refs = []
            for image_info in sheet_images:
                # Create simplified image reference for debug that includes both path and base64
                image_ref = {
                    "filename": image_info["filename"],
                    "index": image_info["index"],
                    "sheet": image_info["sheet"],
                    "position": image_info["position"],
                    "format": image_info["format"],
                    "image_base64": image_info[
                        "image_base64"
                    ],  # Always include base64 data
                }

                # Include path if available (for debugging/logging)
                if "path" in image_info and os.path.exists(image_info["path"]):
                    image_ref["path"] = image_info["path"]

                sheet_image_refs.append(image_ref)

            if sheet_image_refs:
                image_refs[sheet_name] = sheet_image_refs

        # Add image references to debug data
        debug_data["__debug_image_references"] = image_refs

        # Add image extraction summary
        debug_data["__debug_image_summary"] = _create_image_summary(images)

    # Add metadata
    debug_data["__debug_metadata"] = {
        "timestamp": datetime.datetime.now().isoformat(),
        "temp_directory": temp_dir,
        "base_filename": base_filename,
        "enhanced_features": {
            "image_position_extraction": True,
            "position_based_matching": True,
            "format_detection": True,
            "simplified_image_paths": True,
            "base64_embedding": True,  # Add flag for base64 embedding
        },
    }

    # Get session ID from headers or generate a new one
    session_id = request.headers.get("X-Session-ID")
    if not session_id:
        session_id = str(uuid.uuid4())
        logger.info(f"Generated new session ID: {session_id}")
    else:
        logger.info(f"Using provided session ID: {session_id}")

    # Save data to JSON file using the temp_manager with the DEBUG file type
    temp_manager = TempFileManager()
    debug_file_path = temp_manager.save_file_to_temp(
        temp_dir,
        f"{session_id}_debug_data.json",
        json.dumps(debug_data, indent=2, default=str),
        temp_manager.FILE_TYPE_DEBUG,
    )

    logger.info(f"Development mode: Saved enhanced debug data to {debug_file_path}")
    if images:
        total_images = sum(len(sheet_images) for sheet_images in images.values())
        logger.info(
            f"Development mode: Saved {total_images} extracted images with base64 data"
        )

    return debug_file_path


# Google Cloud Function entry point
@functions_framework.http
def excel_pptx_merger(request):
    """Google Cloud Function entry point with enhanced image storage support.

    This function handles HTTP requests to the Cloud Function, supporting both
    local filesystem and Google Cloud Storage for temporary file storage.

    Args:
        request: Flask request object from Cloud Functions

    Returns:
        Flask response with the merged PowerPoint file or error information
    """
    # Setup logging for Cloud Functions environment
    setup_logging()

    # Check if this is a health check request
    if request.path == "/health" or request.path == "/api/v1/health":
        return health()[0]

    # Check authentication
    if not authenticate_request():
        error_response, status_code = create_error_response(
            AuthenticationError("Invalid API key"), 401
        )
        return error_response, status_code

    # Route request to the appropriate endpoint based on path
    if request.method == "POST":
        try:
            # Extract endpoint - requires excel_file and sheet_names
            if request.path == "/api/v1/extract":
                return extract_data_endpoint()

            # Update endpoint - handle before general file validation
            elif request.path == "/api/v1/update":
                return update_excel_file()

            # Check if files were uploaded for other endpoints
            if not request.files:
                # Cloud Functions might receive files differently
                return (
                    jsonify(
                        create_error_response(
                            ValidationError("No files were uploaded"), 400
                        )[0]
                    ),
                    400,
                )

            # Merge and Preview endpoints - require both excel_file and pptx_file
            if request.path in ["/api/v1/merge", "/api/v1/preview"]:
                if (
                    "excel_file" not in request.files
                    or "pptx_file" not in request.files
                ):
                    return (
                        jsonify(
                            create_error_response(
                                ValidationError(
                                    "Both 'excel_file' and 'pptx_file' are required"
                                ),
                                400,
                            )[0]
                        ),
                        400,
                    )

                # Process the request using the appropriate endpoint handler
                if request.path == "/api/v1/merge":
                    return merge_files()
                elif request.path == "/api/v1/preview":
                    return preview_merge()

            # Config endpoint
            elif request.path == "/api/v1/config":
                return manage_config()

            # Stats endpoint
            elif request.path == "/api/v1/stats":
                return get_stats()

            # Unknown endpoint
            else:
                return (
                    jsonify(
                        create_error_response(
                            ValidationError(f"Unknown endpoint: {request.path}"), 404
                        )[0]
                    ),
                    404,
                )

        except Exception as e:
            logger.exception("Error processing Cloud Function request")
            return jsonify(create_error_response(e, 500)[0]), 500
    elif request.method == "GET":
        # Handle GET requests for config and stats
        if request.path == "/api/v1/config":
            return manage_config()
        elif request.path == "/api/v1/stats":
            return get_stats()
        else:
            return (
                jsonify(
                    create_error_response(
                        ValidationError(
                            f"GET method not supported for endpoint: {request.path}"
                        ),
                        405,
                    )[0]
                ),
                405,
            )
    else:
        return (
            jsonify(
                create_error_response(
                    ValidationError("Only POST and GET methods are supported"), 405
                )[0]
            ),
            405,
        )


# CLI interface
@click.group()
def cli():
    """Excel to PowerPoint Merger CLI with enhanced image support."""
    setup_logging()


@cli.command("merge")
@click.option(
    "--excel-file",
    "-e",
    required=True,
    type=click.Path(exists=True),
    help="Path to Excel file",
)
@click.option(
    "--pptx-file",
    "-p",
    required=True,
    type=click.Path(exists=True),
    help="Path to PowerPoint template",
)
@click.option("--output-file", "-o", required=False, help="Output file name")
@click.option(
    "--config-file",
    "-c",
    required=False,
    type=click.Path(exists=True),
    help="Path to extraction configuration JSON file",
)
@click.option(
    "--debug-images/--no-debug-images", default=False, help="Save debug images"
)
def merge_cli(
    excel_file, pptx_file, output_file=None, config_file=None, debug_images=False
):
    """Merge Excel data into PowerPoint template."""
    try:
        # Load extraction configuration if provided
        extraction_config = {}
        if config_file:
            try:
                with open(config_file, "r") as f:
                    extraction_config = json.load(f)
            except json.JSONDecodeError as e:
                click.echo(f"Error parsing config file: {e}", err=True)
                return 1
            except Exception as e:
                click.echo(f"Error reading config file: {e}", err=True)
                return 1
        else:
            # Use default configuration when no config file is provided
            extraction_config = config_manager.get_default_config()

        # Initialize temp file manager
        temp_file_config = extraction_config.get("global_settings", {}).get(
            "temp_file_cleanup", {}
        )
        temp_manager = TempFileManager(temp_file_config)

        # Create a session directory
        session_id = str(uuid.uuid4())
        temp_dir = temp_manager.get_session_directory(session_id)

        # Process Excel file
        excel_processor = ExcelProcessor(excel_file)
        try:
            # Extract data from Excel
            try:
                extracted_data = excel_processor.extract_data(
                    extraction_config.get("global_settings", {}),
                    extraction_config.get("sheet_configs", {}),
                )

                # Extract images with enhanced position information
                images = None
                if (
                    extraction_config.get("global_settings", {})
                    .get("image_extraction", {})
                    .get("enabled", True)
                ):
                    images = excel_processor.extract_images(temp_dir)

                    # Log image extraction summary
                    image_summary = excel_processor.get_image_summary(images)
                    logger.info(f"Image extraction summary: {image_summary}")

            except Exception as e:
                logger.error(f"Failed to extract data from Excel: {e}")
                click.echo(f"Error: Failed to extract data from Excel: {e}", err=True)
                return 1
        finally:
            excel_processor.close()

        # Process PowerPoint template
        pptx_processor = PowerPointProcessor(pptx_file)
        try:
            # Set output file name if not provided
            if not output_file:
                output_file = f"merged_{os.path.basename(pptx_file)}"

            # Merge data into PowerPoint
            merged_file_path = pptx_processor.merge_data(
                extracted_data, output_file, images
            )

            click.echo(
                f"Successfully merged data into PowerPoint template: {merged_file_path}"
            )
            return 0
        finally:
            pptx_processor.close()

    except ValidationError as e:
        click.echo(f"Error: {e}", err=True)
        if debug_images:
            # Save debug images if requested
            click.echo("Saving debug images...")
        return 1
    except Exception as e:
        click.echo(f"Error: {e}", err=True)
        return 1


@cli.command()
@click.option("--host", default="0.0.0.0", help="Host to bind to")
@click.option("--port", default=5000, help="Port to bind to")
@click.option("--debug", is_flag=True, help="Enable debug mode")
def serve(host: str, port: int, debug: bool) -> None:
    """Start the Flask development server."""
    setup_logging()
    logger.info(f"Starting Excel to PowerPoint Merger server on {host}:{port}")
    logger.info("Enhanced features: Image position extraction, position-based matching")

    app.run(
        host=host, port=port, debug=debug or app_config.get("development_mode", False)
    )


if __name__ == "__main__":
    cli()
