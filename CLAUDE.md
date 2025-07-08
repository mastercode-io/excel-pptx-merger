# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Development Commands

### Environment Setup
```bash
# Automated setup (recommended)
python scripts/setup_dev.py

# Manual setup with uv
uv venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
uv pip install -e ".[dev]"
```

### Running the Application
```bash
# Development server with debug mode
python scripts/run_local_server.py --debug

# Direct execution
uv run python -m src.main serve --debug

# CLI merge command
uv run python -m src.main merge -e path/to/excel.xlsx -p path/to/template.pptx
```

### Testing
```bash
# Run all tests
python -m pytest

# Run specific test file
python -m pytest tests/test_excel_processor.py

# Run with coverage
python -m pytest --cov=src --cov-report=term-missing --cov-report=html
```

### Code Quality
```bash
# Format code
black src/ tests/

# Lint code
flake8 src/ tests/

# Type checking
mypy src/
```

### Deployment
```bash
# Deploy to Google Cloud
python scripts/deploy_gcp.py --project your-project-id

# Docker deployment
docker build -f docker/Dockerfile -t excel-pptx-merger:latest .
docker run -p 5000:5000 excel-pptx-merger:latest
```

## Architecture Overview

### Core Components
- **src/main.py**: Flask API server and CLI entry point with Google Cloud Function support
- **src/excel_processor.py**: Excel data extraction with image handling
- **src/pptx_processor.py**: PowerPoint template processing and merge field replacement
- **src/config_manager.py**: Configuration management for extraction rules
- **src/temp_file_manager.py**: Temporary file handling with cleanup

### Key Features
- **Dual Processing Modes**: Memory-only processing for Cloud Functions or file-based for local development
- **Enhanced Image Extraction**: Extracts images from Excel with position information and base64 encoding
- **Flexible Configuration**: JSON-based extraction rules with support for key-value pairs and tables
- **Multi-Environment Support**: Development, testing, and production configurations

### Data Flow
1. Excel files processed using pandas/openpyxl with configurable extraction rules
2. Images extracted with position metadata and embedded as base64
3. PowerPoint templates processed using python-pptx library
4. Jinja-style merge fields ({{field_name}}) replaced with extracted data
5. Images inserted into placeholders with aspect ratio preservation

### Configuration System
- **config/default_config.json**: Default extraction configuration
- **Environment files**: development.env, production.env, testing.env
- **Runtime configuration**: Passed via API requests or CLI parameters

### API Endpoints
- **POST /api/v1/merge**: Main processing endpoint
- **POST /api/v1/preview**: Preview extraction without merging
- **POST /api/v1/extract**: Extract data from single Excel sheet
- **POST /api/v1/update**: Update Excel file with data (supports dual payload modes)
- **GET /api/v1/config**: Get default configuration
- **GET /api/v1/health**: Health check with feature status

### Testing Structure
- **Unit tests**: Individual component testing
- **Integration tests**: End-to-end workflow testing
- **Fixtures**: Sample Excel/PowerPoint files in tests/fixtures/

### File Organization
- **src/**: Main application code
- **src/utils/**: Utility modules (exceptions, validation, file handling)
- **config/**: Configuration files and environment settings
- **scripts/**: Development and deployment scripts
- **tests/**: Test files and fixtures
- **docker/**: Docker configuration files

## Development Notes

### Memory vs File-Based Processing
The application supports both memory-only processing (for Cloud Functions) and file-based processing (for local development). This is controlled by the `save_files` configuration option.

### Dual Payload Mode Support
The `/api/v1/update` endpoint supports two payload formats for maximum client compatibility:

#### **Multipart Mode (Standard Clients)**
- **Content-Type**: `multipart/form-data`
- **Excel File**: Binary file upload
- **Data**: JSON strings in form fields
- **Use Case**: Postman, web browsers, standard HTTP clients

#### **JSON Mode (CRM/Automation Systems)**
- **Content-Type**: `application/json` or `text/plain`
- **Excel File**: Base64-encoded string in JSON
- **Data**: Direct JSON objects
- **Use Case**: Deluge CRM, automation systems, webhooks

The endpoint automatically detects the payload format and processes both modes through a unified processing path.

### CRM System Compatibility
Special handling for CRM and automation systems that may:
- Send JSON data with incorrect Content-Type headers (`text/plain` instead of `application/json`)
- Include large base64-encoded images in payloads
- Add additional metadata or tracking information

The system includes enhanced Content-Type detection and robust JSON parsing to handle these scenarios.

### Image Handling
Images are extracted from Excel files with position information and converted to base64 for embedding in JSON responses. The system supports position-based matching between Excel images and PowerPoint placeholders.

### Error Handling
Uses custom exception classes in src/utils/exceptions.py for structured error handling across the application. Enhanced 413 error handling provides detailed diagnostics for payload size issues.

### Debug Mode
Development mode saves debug information including extracted data and copies of processed files to the debug directory for troubleshooting. Includes enhanced logging for payload type detection and processing modes.

### Testing Philosophy
- **User handles comprehensive testing**: The user will run merge tests with real data as it's faster and more representative of actual usage scenarios
- **Claude runs basic verification only**: Only run basic tests (like unit tests or simple integration tests) to verify code functionality and catch obvious regressions
- **Real-world validation**: Complex scenarios, edge cases, and performance testing are best handled by the user with actual production data