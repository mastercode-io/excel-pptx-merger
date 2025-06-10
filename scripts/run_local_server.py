#!/usr/bin/env python3
"""Local Flask server for testing and development."""

import os
import sys
import click
import logging
from pathlib import Path

# Add src to path
script_dir = Path(__file__).parent
project_root = script_dir.parent
sys.path.insert(0, str(project_root / 'src'))

from src.main import app, setup_logging
from src.config_manager import ConfigManager

@click.command()
@click.option('--host', default='0.0.0.0', help='Host to bind to')
@click.option('--port', default=5000, type=int, help='Port to bind to')
@click.option('--debug', is_flag=True, help='Enable debug mode')
@click.option('--env', default='development', help='Environment (development, testing, production)')
@click.option('--reload', is_flag=True, help='Enable auto-reload on file changes')
@click.option('--log-level', default='INFO', help='Log level (DEBUG, INFO, WARNING, ERROR)')
def run_server(host, port, debug, env, reload, log_level):
    """Run the Excel to PowerPoint Merger Flask server locally."""
    
    # Set environment
    os.environ['ENVIRONMENT'] = env
    os.environ['DEVELOPMENT_MODE'] = str(debug).lower()
    os.environ['LOG_LEVEL'] = log_level.upper()
    
    # Setup logging
    setup_logging()
    logger = logging.getLogger(__name__)
    
    # Load environment-specific configuration
    config_manager = ConfigManager()
    app_config = config_manager.get_app_config()
    
    logger.info(f"Starting Excel to PowerPoint Merger server")
    logger.info(f"Environment: {env}")
    logger.info(f"Debug mode: {debug}")
    logger.info(f"Host: {host}")
    logger.info(f"Port: {port}")
    logger.info(f"Log level: {log_level}")
    
    # Configure Flask app
    app.config.update({
        'DEBUG': debug,
        'TESTING': env == 'testing',
        'ENV': env
    })
    
    if debug:
        logger.info("Debug mode enabled - detailed error messages and auto-reload")
        logger.info("Available endpoints:")
        logger.info("  GET  /api/v1/health      - Health check")
        logger.info("  POST /api/v1/merge       - Merge Excel and PowerPoint files")
        logger.info("  POST /api/v1/preview     - Preview merge without processing")
        logger.info("  GET  /api/v1/config      - Get default configuration")
        logger.info("  POST /api/v1/config      - Validate configuration")
        logger.info("  GET  /api/v1/stats       - Get system statistics")
    
    try:
        # Run the Flask development server
        app.run(
            host=host,
            port=port,
            debug=debug,
            use_reloader=reload,
            threaded=True
        )
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server error: {e}")
        sys.exit(1)

if __name__ == '__main__':
    run_server()