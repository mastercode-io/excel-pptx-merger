"""Main API handler with Flask and Google Cloud Function support - Enhanced with image position support."""

import json
import logging
import os
import traceback
import datetime
import copy
from typing import Any, Dict, List, Optional, Tuple, Union
from flask import Flask, request, jsonify, send_file
from werkzeug.exceptions import RequestEntityTooLarge
import click
import functions_framework

from .config_manager import ConfigManager
from .excel_processor import ExcelProcessor
from .pptx_processor import PowerPointProcessor
from .temp_file_manager import TempFileManager
from .utils.exceptions import (
    ExcelPptxMergerError, ValidationError, FileProcessingError,
    ConfigurationError, APIError, AuthenticationError
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
app.config['MAX_CONTENT_LENGTH'] = app_config['max_file_size_mb'] * 1024 * 1024


def setup_logging() -> None:
    """Setup logging configuration."""
    log_level = app_config.get('log_level', 'INFO').upper()
    logging.getLogger().setLevel(getattr(logging, log_level, logging.INFO))

    # Set up format
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )

    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logging.getLogger().addHandler(console_handler)


def authenticate_request() -> bool:
    """Authenticate API request."""
    # Skip authentication in development mode
    if app_config.get('development_mode', False):
        logger.debug("Authentication bypassed in development mode")
        return True

    api_key = app_config.get('api_key')
    if not api_key:
        return True  # No authentication required if no key configured

    # Check Authorization header
    auth_header = request.headers.get('Authorization')
    if auth_header and auth_header.startswith('Bearer '):
        token = auth_header.split(' ', 1)[1]
        return token == api_key

    # Check query parameter
    request_key = request.args.get('api_key') or request.form.get('api_key')
    return request_key == api_key


def create_error_response(error: Exception, status_code: int = 500) -> Tuple[Dict[str, Any], int]:
    """Create standardized error response."""
    error_response = {
        'success': False,
        'error': {
            'type': type(error).__name__,
            'message': str(error),
            'code': status_code
        }
    }

    # Add error code if available
    if hasattr(error, 'error_code') and error.error_code:
        error_response['error']['error_code'] = error.error_code

    # Add traceback in development mode
    if app_config.get('development_mode', False):
        error_response['error']['traceback'] = traceback.format_exc()

    logger.error(f"API Error ({status_code}): {error}")
    return error_response, status_code


@app.errorhandler(RequestEntityTooLarge)
def handle_file_too_large(error):
    """Handle file size too large error."""
    return create_error_response(
        ValidationError(f"File size exceeds maximum allowed size of {app_config['max_file_size_mb']}MB"),
        413
    )


@app.before_request
def before_request():
    """Pre-request authentication and validation."""
    # Skip authentication for health check
    if request.endpoint == 'health':
        return

    # Authenticate request
    if not authenticate_request():
        error_response, status_code = create_error_response(
            AuthenticationError("Invalid API key"), 401
        )
        return jsonify(error_response), status_code


@app.route('/api/v1/health', methods=['GET'])
def health() -> Tuple[Dict[str, Any], int]:
    """Health check endpoint."""
    try:
        health_info = {
            'success': True,
            'status': 'healthy',
            'version': '0.1.0',
            'environment': os.getenv('ENVIRONMENT', 'development'),
            'timestamp': None,
            'services': {
                'config_manager': True,
                'temp_file_manager': True
            },
            'features': {
                'enhanced_image_extraction': True,
                'position_based_image_matching': True,
                'debug_mode': app_config.get('development_mode', False)
            }
        }

        # Add timestamp
        from datetime import datetime
        health_info['timestamp'] = datetime.utcnow().isoformat() + 'Z'

        return health_info, 200

    except Exception as e:
        return create_error_response(e, 500)


@app.route('/api/v1/merge', methods=['POST'])
def merge_files() -> Union[Tuple[Dict[str, Any], int], Any]:
    """Main file processing endpoint with enhanced image support."""
    temp_manager = None
    temp_dir = None

    try:
        # Validate request
        if 'excel_file' not in request.files or 'pptx_file' not in request.files:
            raise ValidationError("Both 'excel_file' and 'pptx_file' are required")

        excel_file = request.files['excel_file']
        pptx_file = request.files['pptx_file']

        if excel_file.filename == '' or pptx_file.filename == '':
            raise ValidationError("File names cannot be empty")

        # Validate file extensions
        allowed_extensions = app_config['allowed_extensions']
        if not excel_file.filename.lower().endswith('.xlsx'):
            raise ValidationError("Excel file must have .xlsx extension")

        if not pptx_file.filename.lower().endswith('.pptx'):
            raise ValidationError("PowerPoint file must have .pptx extension")

        # Get configuration
        config_data = request.form.get('config')
        if config_data:
            try:
                extraction_config = json.loads(config_data)
            except json.JSONDecodeError as e:
                raise ValidationError(f"Invalid JSON configuration: {e}")
        else:
            extraction_config = config_manager.get_default_config()

        # Validate configuration
        config_manager.validate_runtime_config(extraction_config)

        # Initialize temp file manager
        temp_file_config = extraction_config.get('global_settings', {}).get('temp_file_cleanup', {})
        temp_manager = TempFileManager(temp_file_config)

        # Create temporary directory
        temp_dir = temp_manager.create_temp_directory()

        # Save uploaded files
        excel_path = temp_manager.save_file_to_temp(
            temp_dir, f"input_{excel_file.filename}", excel_file
        )
        pptx_path = temp_manager.save_file_to_temp(
            temp_dir, f"template_{pptx_file.filename}", pptx_file
        )

        logger.info(f"Processing files: {excel_file.filename} + {pptx_file.filename}")

        # Process Excel file
        excel_processor = ExcelProcessor(excel_path)
        try:
            # Extract data according to configuration
            extracted_data = excel_processor.extract_data(
                extraction_config.get('sheet_configs', {})
            )

            # Extract images with enhanced position information
            images = None
            if extraction_config.get('global_settings', {}).get('image_extraction', {}).get('enabled', True):
                # Pass the global settings to extract_images
                images = excel_processor.extract_images(extraction_config.get('global_settings', {}))

                # Link images to the image search table
                if images:
                    extracted_data = excel_processor.link_images_to_table(extracted_data, images)
                    
                    # Log image extraction summary
                    image_summary = excel_processor.get_image_summary(images)
                    logger.info(f"Image extraction summary: {image_summary}")

            # Save debug information in development mode
            base_filename = os.path.splitext(excel_file.filename)[0]
            debug_file_path = save_debug_info(extracted_data, images, temp_dir, base_filename)

        finally:
            excel_processor.close()

        # Process PowerPoint template
        pptx_processor = PowerPointProcessor(pptx_path)
        try:
            # Generate output filename
            output_filename = f"merged_{pptx_file.filename}"
            output_path = os.path.join(temp_dir, output_filename)

            # Merge data into template with enhanced image support
            merged_file_path = pptx_processor.merge_data(extracted_data, output_path, images)

            # In development mode, also save a copy to tests/fixtures
            if app_config.get('development_mode', False):
                fixtures_dir = os.path.join(os.getcwd(), 'tests', 'fixtures')
                os.makedirs(fixtures_dir, exist_ok=True)
                fixtures_output_path = os.path.join(fixtures_dir, output_filename)
                import shutil
                shutil.copy2(merged_file_path, fixtures_output_path)
                logger.info(f"Development mode: Saved copy of merged file to {fixtures_output_path}")

            # Clean up images after successful merge if configured
            if images:
                excel_processor.cleanup_images(images, extraction_config.get('global_settings', {}))

            # Return the merged file
            return send_file(
                merged_file_path,
                as_attachment=True,
                download_name=output_filename,
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
            )

        finally:
            pptx_processor.close()

    except ValidationError as e:
        if temp_manager and temp_dir:
            temp_manager.mark_error(temp_dir)
        return jsonify(create_error_response(e, 400)[0]), 400

    except (FileProcessingError, ConfigurationError) as e:
        if temp_manager and temp_dir:
            temp_manager.mark_error(temp_dir)
        return jsonify(create_error_response(e, 422)[0]), 422

    except Exception as e:
        if temp_manager and temp_dir:
            temp_manager.mark_error(temp_dir)
        return jsonify(create_error_response(e, 500)[0]), 500

    finally:
        # Schedule cleanup
        if temp_manager and temp_dir:
            temp_manager.schedule_cleanup(temp_dir)


@app.route('/api/v1/preview', methods=['POST'])
def preview_merge() -> Tuple[Dict[str, Any], int]:
    """Preview merge without performing actual merge - Enhanced with image analysis."""
    temp_manager = None
    temp_dir = None

    try:
        # Validate request
        if 'excel_file' not in request.files or 'pptx_file' not in request.files:
            raise ValidationError("Both 'excel_file' and 'pptx_file' are required")

        excel_file = request.files['excel_file']
        pptx_file = request.files['pptx_file']

        # Get configuration
        config_data = request.form.get('config')
        if config_data:
            extraction_config = json.loads(config_data)
        else:
            extraction_config = config_manager.get_default_config()

        # Initialize temp file manager
        temp_file_config = extraction_config.get('global_settings', {}).get('temp_file_cleanup', {})
        temp_manager = TempFileManager(temp_file_config)
        temp_dir = temp_manager.create_temp_directory()

        # Save files
        excel_path = temp_manager.save_file_to_temp(temp_dir, excel_file.filename, excel_file)
        pptx_path = temp_manager.save_file_to_temp(temp_dir, pptx_file.filename, pptx_file)

        # Process Excel file
        excel_processor = ExcelProcessor(excel_path)
        try:
            extracted_data = excel_processor.extract_data(
                extraction_config.get('sheet_configs', {})
            )

            # Extract images with position information
            images = None
            if extraction_config.get('global_settings', {}).get('image_extraction', {}).get('enabled', True):
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
            'success': True,
            'preview': {
                'extracted_data': extracted_data,
                'template_info': template_info,
                'merge_preview': preview_info,
                'image_analysis': {
                    'extracted_images': _create_image_summary(images) if images else {},
                    'image_requirements': image_requirements,
                    'matching_analysis': _analyze_image_matching(images, preview_info.get('image_placeholders', []))
                },
                'configuration_used': extraction_config
            }
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
        'total_images': 0,
        'sheets': {},
        'position_info': {
            'images_with_position': 0,
            'images_without_position': 0,
            'anchor_types': {}
        }
    }

    for sheet_name, sheet_images in images.items():
        sheet_summary = {
            'count': len(sheet_images),
            'images': []
        }

        for image_info in sheet_images:
            image_summary = {
                'filename': image_info.get('filename'),
                'index': image_info.get('index'),
                'format': image_info.get('format'),
                'size': image_info.get('size'),
                'position': image_info.get('position', {})
            }

            # Count position statistics
            position = image_info.get('position', {})
            if position.get('estimated_cell'):
                summary['position_info']['images_with_position'] += 1
                anchor_type = position.get('anchor_type', 'unknown')
                summary['position_info']['anchor_types'][anchor_type] = \
                    summary['position_info']['anchor_types'].get(anchor_type, 0) + 1
            else:
                summary['position_info']['images_without_position'] += 1

            sheet_summary['images'].append(image_summary)

        summary['sheets'][sheet_name] = sheet_summary
        summary['total_images'] += len(sheet_images)

    return summary


def _analyze_image_matching(images: Optional[Dict[str, List[Dict[str, Any]]]],
                           placeholders: List[str]) -> Dict[str, Any]:
    """Analyze potential image-placeholder matching."""
    analysis = {
        'total_placeholders': len(placeholders),
        'potential_matches': {},
        'unmatched_placeholders': [],
        'unmatched_images': [],
        'recommendations': []
    }

    if not images or not placeholders:
        analysis['unmatched_placeholders'] = placeholders.copy()
        if images:
            all_images = []
            for sheet_images in images.values():
                all_images.extend(sheet_images)
            analysis['unmatched_images'] = [img.get('filename') for img in all_images]
        return analysis

    # Create flat list of all images with their info
    all_images = []
    for sheet_name, sheet_images in images.items():
        for image_info in sheet_images:
            all_images.append({
                'sheet': sheet_name,
                'info': image_info
            })

    matched_images = set()

    # Try to match placeholders with images
    for placeholder in placeholders:
        best_match = None
        match_confidence = 0

        for idx, image_entry in enumerate(all_images):
            if idx in matched_images:
                continue

            confidence = _calculate_match_confidence(placeholder, image_entry['info'])
            if confidence > match_confidence:
                match_confidence = confidence
                best_match = {
                    'image_index': idx,
                    'confidence': confidence,
                    'image_info': image_entry
                }

        if best_match and match_confidence > 0.3:  # Threshold for reasonable match
            analysis['potential_matches'][placeholder] = best_match
            matched_images.add(best_match['image_index'])
        else:
            analysis['unmatched_placeholders'].append(placeholder)

    # Find unmatched images
    for idx, image_entry in enumerate(all_images):
        if idx not in matched_images:
            analysis['unmatched_images'].append(image_entry['info'].get('filename'))

    # Generate recommendations
    analysis['recommendations'] = _generate_matching_recommendations(analysis)

    return analysis


def _calculate_match_confidence(placeholder: str, image_info: Dict[str, Any]) -> float:
    """Calculate confidence score for placeholder-image matching."""
    confidence = 0.0
    placeholder_lower = placeholder.lower()

    # Position-based matching (highest confidence)
    position = image_info.get('position', {})
    if position.get('estimated_cell'):
        cell_ref = position['estimated_cell'].lower()
        if cell_ref in placeholder_lower or placeholder_lower.endswith(cell_ref):
            confidence += 0.8

    # Index-based matching
    import re
    placeholder_numbers = re.findall(r'\d+', placeholder_lower)
    image_index = image_info.get('index', 0)

    if placeholder_numbers:
        for num_str in placeholder_numbers:
            try:
                num = int(num_str)
                if num == image_index or num == image_index - 1:  # 0-based or 1-based indexing
                    confidence += 0.6
                    break
            except ValueError:
                continue

    # Keyword matching
    keywords = ['image', 'img', 'picture', 'photo']
    for keyword in keywords:
        if keyword in placeholder_lower:
            confidence += 0.3
            break

    # Sheet name matching
    sheet_name = image_info.get('sheet', '').lower()
    if sheet_name and sheet_name.replace(' ', '_') in placeholder_lower:
        confidence += 0.4

    return min(confidence, 1.0)  # Cap at 1.0


def _generate_matching_recommendations(analysis: Dict[str, Any]) -> List[str]:
    """Generate recommendations for improving image matching."""
    recommendations = []

    unmatched_count = len(analysis['unmatched_placeholders'])
    unmatched_images_count = len(analysis['unmatched_images'])

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

    if analysis['potential_matches']:
        low_confidence_matches = [
            placeholder for placeholder, match in analysis['potential_matches'].items()
            if match['confidence'] < 0.6
        ]

        if low_confidence_matches:
            recommendations.append(
                f"Consider improving placeholder names for better matching: {', '.join(low_confidence_matches)}"
            )

    if not recommendations:
        recommendations.append("Image matching analysis looks good! All placeholders have potential matches.")

    return recommendations


@app.route('/api/v1/config', methods=['GET', 'POST'])
def manage_config() -> Tuple[Dict[str, Any], int]:
    """Manage configuration with enhanced image extraction settings."""
    try:
        if request.method == 'GET':
            # Return default configuration with enhanced image settings
            config = config_manager.get_default_config()
            return {
                'success': True,
                'config': config,
                'features': {
                    'enhanced_image_extraction': True,
                    'position_based_matching': True,
                    'supported_image_formats': ['png', 'jpg', 'jpeg', 'gif', 'webp']
                }
            }, 200

        elif request.method == 'POST':
            # Validate and save configuration
            config_data = request.get_json()
            if not config_data:
                raise ValidationError("Configuration data is required")

            config_manager.validate_runtime_config(config_data)

            # For now, just validate - future versions could store custom configs
            return {
                'success': True,
                'message': 'Configuration validated successfully',
                'config': config_data
            }, 200

    except Exception as e:
        return create_error_response(e, 400)


@app.route('/api/v1/stats', methods=['GET'])
def get_stats() -> Tuple[Dict[str, Any], int]:
    """Get system statistics with enhanced feature information."""
    try:
        # This could be expanded to include more detailed stats
        stats = {
            'success': True,
            'stats': {
                'app_config': {
                    'max_file_size_mb': app_config['max_file_size_mb'],
                    'allowed_extensions': app_config['allowed_extensions'],
                    'development_mode': app_config['development_mode']
                },
                'features': {
                    'enhanced_image_extraction': True,
                    'position_based_image_matching': True,
                    'image_format_detection': True,
                    'debug_mode': app_config.get('development_mode', False)
                },
                'runtime': {
                    'python_version': os.sys.version,
                    'environment': os.getenv('ENVIRONMENT', 'development')
                }
            }
        }

        return stats, 200

    except Exception as e:
        return create_error_response(e, 500)


def save_debug_info(extracted_data, images, temp_dir, base_filename):
    """Save enhanced debug information with image positions to files for development purposes."""
    if not app_config.get('development_mode', False):
        return None

    # Create fixtures directory if it doesn't exist
    fixtures_dir = os.path.join(os.getcwd(), 'tests', 'fixtures')
    os.makedirs(fixtures_dir, exist_ok=True)

    # Create images directory if it doesn't exist
    images_dir = os.path.join(fixtures_dir, 'images')
    os.makedirs(images_dir, exist_ok=True)

    # Create a deep copy of the extracted data to ensure we don't modify the original
    debug_data = copy.deepcopy(extracted_data)

    # Process image structure for debug data
    if images:
        image_refs = {}
        for sheet_name, sheet_images in images.items():
            sheet_image_refs = []
            for image_info in sheet_images:
                if os.path.exists(image_info['path']):
                    # Get image filename
                    image_filename = os.path.basename(image_info['path'])
                    
                    # Create reference to the image path
                    # No need to copy the image as it's already in the right location
                    # if we're using the development mode directory
                    
                    # Create simplified image reference for debug
                    image_ref = {
                        'path': image_info['path'],
                        'filename': image_info['filename'],
                        'index': image_info['index'],
                        'sheet': image_info['sheet'],
                        'position': image_info['position'],
                        'size': image_info['size'],
                        'format': image_info['format']
                    }
                    sheet_image_refs.append(image_ref)

            if sheet_image_refs:
                image_refs[sheet_name] = sheet_image_refs

        # Add image references to debug data
        debug_data['__debug_image_references'] = image_refs

        # Add image extraction summary
        debug_data['__debug_image_summary'] = _create_image_summary(images)

    # Add metadata
    debug_data['__debug_metadata'] = {
        'timestamp': datetime.datetime.now().isoformat(),
        'temp_directory': temp_dir,
        'base_filename': base_filename,
        'enhanced_features': {
            'image_position_extraction': True,
            'position_based_matching': True,
            'format_detection': True,
            'simplified_image_paths': True
        }
    }
    
    # Save data to JSON file
    debug_file_path = os.path.join(fixtures_dir, f"{base_filename}_debug_data.json")
    with open(debug_file_path, 'w') as f:
        json.dump(debug_data, f, indent=2, default=str)

    logger.info(f"Development mode: Saved enhanced debug data to {debug_file_path}")
    if images:
        total_images = sum(len(sheet_images) for sheet_images in images.values())
        logger.info(f"Development mode: Saved {total_images} extracted images to {images_dir}")
    
    return debug_file_path


# Google Cloud Function entry point
@functions_framework.http
def excel_pptx_merger(request):
    """Google Cloud Function entry point."""
    with app.request_context(request.environ):
        try:
            # Route the request based on path
            path = request.path.rstrip('/')
            method = request.method

            if path == '/api/v1/health' and method == 'GET':
                response_data, status_code = health()
            elif path == '/api/v1/merge' and method == 'POST':
                return merge_files()  # This returns a file response
            elif path == '/api/v1/preview' and method == 'POST':
                response_data, status_code = preview_merge()
            elif path == '/api/v1/config' and method in ['GET', 'POST']:
                response_data, status_code = manage_config()
            elif path == '/api/v1/stats' and method == 'GET':
                response_data, status_code = get_stats()
            else:
                response_data, status_code = create_error_response(
                    APIError(f"Endpoint not found: {method} {path}"), 404
                )

            return jsonify(response_data), status_code

        except Exception as e:
            response_data, status_code = create_error_response(e, 500)
            return jsonify(response_data), status_code


# CLI interface
@click.group()
def cli():
    """Excel to PowerPoint Merger CLI with enhanced image support."""
    setup_logging()


@cli.command()
@click.option('--host', default='0.0.0.0', help='Host to bind to')
@click.option('--port', default=5000, help='Port to bind to')
@click.option('--debug', is_flag=True, help='Enable debug mode')
def serve(host: str, port: int, debug: bool) -> None:
    """Start the Flask development server."""
    setup_logging()
    logger.info(f"Starting Excel to PowerPoint Merger server on {host}:{port}")
    logger.info("Enhanced features: Image position extraction, position-based matching")

    app.run(
        host=host,
        port=port,
        debug=debug or app_config.get('development_mode', False)
    )


@cli.command()
@click.argument('excel_file', type=click.Path(exists=True))
@click.argument('pptx_file', type=click.Path(exists=True))
@click.argument('output_file', type=click.Path())
@click.option('--config', type=click.Path(exists=True), help='Configuration file path')
@click.option('--debug-images', is_flag=True, help='Enable detailed image debugging')
def merge(excel_file: str, pptx_file: str, output_file: str, config: Optional[str], debug_images: bool) -> None:
    """Merge Excel data into PowerPoint template via CLI with enhanced image support."""
    try:
        setup_logging()

        # Load configuration
        if config:
            with open(config, 'r') as f:
                extraction_config = json.load(f)
        else:
            extraction_config = config_manager.get_default_config()

        # Initialize temp file manager
        temp_file_config = extraction_config.get('global_settings', {}).get('temp_file_cleanup', {})
        temp_manager = TempFileManager(temp_file_config)

        # Create temporary directory
        with temp_manager.temp_directory() as temp_dir:
            # Process Excel file
            excel_processor = ExcelProcessor(excel_file)
            try:
                extracted_data = excel_processor.extract_data(
                    extraction_config.get('sheet_configs', {})
                )

                # Extract images with enhanced position information
                images = None
                if extraction_config.get('global_settings', {}).get('image_extraction', {}).get('enabled', True):
                    images = excel_processor.extract_images(temp_dir)
                    
                    # Log image extraction summary
                    image_summary = excel_processor.get_image_summary(images)
                    logger.info(f"Image extraction summary: {image_summary}")

            finally:
                excel_processor.close()

            # Process PowerPoint template
            pptx_processor = PowerPointProcessor(pptx_file)
            try:
                merged_file_path = pptx_processor.merge_data(extracted_data, output_file, images)
                click.echo(f"Successfully merged files. Output saved to: {merged_file_path}")

                if debug_images:
                    # Show image matching analysis
                    preview = pptx_processor.preview_merge(extracted_data, images)
                    if 'image_mappings' in preview:
                        click.echo("\n=== Image Mapping Analysis ===")
                        for placeholder, image_path in preview['image_mappings'].items():
                            status = "✓ Matched" if image_path else "✗ Not found"
                            click.echo(f"  {placeholder}: {status}")
                            if image_path:
                                click.echo(f"    -> {os.path.basename(image_path)}")
            finally:
                pptx_processor.close()

    except Exception as e:
        click.echo(f"Error: {e}", err=True)
        if debug_images:
            click.echo(f"Traceback: {traceback.format_exc()}", err=True)
        raise click.Abort()


if __name__ == '__main__':
    cli()
