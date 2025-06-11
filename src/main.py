"""Main API handler with Flask and Google Cloud Function support."""

import json
import logging
import os
import traceback
from typing import Any, Dict, Optional, Tuple, Union
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
    """Main file processing endpoint."""
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
            
            # Extract images if enabled
            images = None
            if extraction_config.get('global_settings', {}).get('image_extraction', {}).get('enabled', True):
                images = excel_processor.extract_images(temp_dir)
            
        finally:
            excel_processor.close()
        
        # Process PowerPoint template
        pptx_processor = PowerPointProcessor(pptx_path)
        try:
            # Generate output filename
            output_filename = f"merged_{pptx_file.filename}"
            output_path = os.path.join(temp_dir, output_filename)
            
            # Merge data into template
            merged_file_path = pptx_processor.merge_data(extracted_data, output_path, images)
            
            # In development mode, also save a copy to tests/fixtures
            if app_config.get('development_mode', False):
                fixtures_dir = os.path.join(os.getcwd(), 'tests', 'fixtures')
                os.makedirs(fixtures_dir, exist_ok=True)
                fixtures_output_path = os.path.join(fixtures_dir, output_filename)
                import shutil
                shutil.copy2(merged_file_path, fixtures_output_path)
                logger.info(f"Development mode: Saved copy of merged file to {fixtures_output_path}")
            
        finally:
            pptx_processor.close()
        
        # Return the merged file
        return send_file(
            merged_file_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
    
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
    """Preview merge without performing actual merge."""
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
        finally:
            excel_processor.close()
        
        # Analyze PowerPoint template
        pptx_processor = PowerPointProcessor(pptx_path)
        try:
            template_info = pptx_processor.validate_template()
            preview_info = pptx_processor.preview_merge(extracted_data)
        finally:
            pptx_processor.close()
        
        # Compile preview response
        preview_response = {
            'success': True,
            'preview': {
                'extracted_data': extracted_data,
                'template_info': template_info,
                'merge_preview': preview_info,
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


@app.route('/api/v1/config', methods=['GET', 'POST'])
def manage_config() -> Tuple[Dict[str, Any], int]:
    """Manage configuration."""
    try:
        if request.method == 'GET':
            # Return default configuration
            config = config_manager.get_default_config()
            return {
                'success': True,
                'config': config
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
    """Get system statistics."""
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
                'runtime': {
                    'python_version': os.sys.version,
                    'environment': os.getenv('ENVIRONMENT', 'development')
                }
            }
        }
        
        return stats, 200
    
    except Exception as e:
        return create_error_response(e, 500)


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
    """Excel to PowerPoint Merger CLI."""
    setup_logging()


@cli.command()
@click.option('--host', default='0.0.0.0', help='Host to bind to')
@click.option('--port', default=5000, help='Port to bind to')
@click.option('--debug', is_flag=True, help='Enable debug mode')
def serve(host: str, port: int, debug: bool) -> None:
    """Start the Flask development server."""
    setup_logging()
    logger.info(f"Starting Excel to PowerPoint Merger server on {host}:{port}")
    
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
def merge(excel_file: str, pptx_file: str, output_file: str, config: Optional[str]) -> None:
    """Merge Excel data into PowerPoint template via CLI."""
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
                
                # Extract images
                images = excel_processor.extract_images(temp_dir)
            finally:
                excel_processor.close()
            
            # Process PowerPoint template
            pptx_processor = PowerPointProcessor(pptx_file)
            try:
                merged_file_path = pptx_processor.merge_data(extracted_data, output_file, images)
                click.echo(f"Successfully merged files. Output saved to: {merged_file_path}")
            finally:
                pptx_processor.close()
    
    except Exception as e:
        click.echo(f"Error: {e}", err=True)
        raise click.Abort()


if __name__ == '__main__':
    cli()