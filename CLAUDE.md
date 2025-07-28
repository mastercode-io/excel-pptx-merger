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
# Deploy to Google Cloud Functions (Europe West 2)
./deploy.sh

# Alternative: Manual deployment with specific settings
gcloud functions deploy excel-pptx-merger \
    --gen2 --runtime=python312 --region=europe-west2 \
    --source=. --entry-point=excel_pptx_merger \
    --trigger-http --allow-unauthenticated \
    --memory=1024MB --timeout=540s \
    --set-env-vars="STORAGE_BACKEND=LOCAL,SAVE_FILES=false,FLASK_DEBUG=false,DEVELOPMENT_MODE=false,LOG_LEVEL=INFO" \
    --set-secrets="GRAPH_CLIENT_ID=excel-pptx-merger-graph-client-id:latest,GRAPH_CLIENT_SECRET=excel-pptx-merger-graph-client-secret:latest,GRAPH_TENANT_ID=excel-pptx-merger-graph-tenant-id:latest"

# Docker deployment (for local development)
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
- **src/job_queue.py**: Async job queue system for long-running operations
- **src/job_handlers.py**: Job handlers with complete code reuse via internal_data parameter
- **src/utils/sharepoint_file_handler.py**: Centralized SharePoint file access for all endpoints
- **src/utils/request_handler.py**: Unified request processing with SharePoint integration
- **src/utils/storage.py**: Storage abstraction layer (Local/GCS) with proper Content-Type handling

### Key Features
- **Async Job Queue System**: Overcomes 10-second timeout limits with background processing and real-time progress tracking
- **Complete Code Reuse**: Job handlers call original endpoints with internal_data parameter (zero duplication)
- **Dual Processing Modes**: Memory-only processing for Cloud Functions or file-based for local development
- **Enhanced Image Extraction**: Extracts images from Excel with position information and base64 encoding
- **Flexible Configuration**: JSON-based extraction rules with support for key-value pairs and tables
- **Multi-Environment Support**: Development, testing, and production configurations
- **Centralized SharePoint Integration**: Unified SharePoint file access across all endpoints via SharePointFileHandler
- **Robust Storage Layer**: Abstracted storage with Local/GCS backends and proper Content-Type handling

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

#### Core Processing Endpoints
- **POST /api/v1/merge**: Main processing endpoint
- **POST /api/v1/preview**: Preview extraction without merging
- **POST /api/v1/extract**: Extract data from single Excel sheet (supports internal_data parameter)
- **POST /api/v1/update**: Update Excel file with data (supports dual payload modes)
- **GET /api/v1/config**: Get default configuration
- **GET /api/v1/health**: Health check with feature status

#### Async Job Queue Endpoints
- **POST /api/v1/jobs/start**: Start async job for any supported endpoint
- **GET /api/v1/jobs/{jobId}/status**: Check job progress and status
- **GET /api/v1/jobs/{jobId}/result**: Get completed job results (with cleanup)
- **GET /api/v1/jobs**: List all jobs (optional, for debugging)
- **DELETE /api/v1/jobs/{jobId}**: Cancel/delete a job
- **GET /api/v1/jobs/stats**: Get job queue statistics

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
- **docs/**: Technical documentation and integration guides

## Development Notes

### Async Job Queue System
The application implements a comprehensive async job queue to overcome Cloud Function timeout limits:

#### **Key Components**
- **src/job_queue.py**: Core job queue with status tracking and cleanup
- **src/job_handlers.py**: Endpoint handlers with complete code reuse via internal_data parameter
- **Job statuses**: pending → running → completed/failed/expired
- **Automatic cleanup**: Jobs expire after 5 minutes, old jobs cleaned up automatically

#### **Code Reuse Strategy**
- **Critical requirement**: NO code duplication between sync and async paths
- **Solution**: Job handlers call original endpoints with `internal_data` parameter
- **extract_data_endpoint**: Modified to accept `internal_data` for direct data passing
- **merge_files & update_excel_file**: Still use MockRequest approach (to be updated)

#### **Production Configuration**
- **Memory-only processing**: `SAVE_FILES=false` for production (fast, no GCS dependencies)
- **File-based processing**: `SAVE_FILES=true` for development/debugging only
- **Storage backends**: Local (production) vs GCS (development)

### Memory vs File-Based Processing
The application supports both memory-only processing (for Cloud Functions) and file-based processing (for local development). This is controlled by the `save_files` configuration option.

**Production (Cloud Functions)**:
- `SAVE_FILES=false` → Memory-only processing
- `STORAGE_BACKEND=LOCAL` → No GCS dependencies
- `DEVELOPMENT_MODE=false` → Minimal logging, no debug files

**Development (Local)**:
- `SAVE_FILES=true` → Files saved for debugging
- `STORAGE_BACKEND=GCS` → Uses Google Cloud Storage
- `DEVELOPMENT_MODE=true` → Enhanced logging, debug files

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

## Related Documentation Files

### Technical Implementation Guides
- **docs/gcloud-job-queue-implementation.md**: Complete technical implementation details of the async job queue system
- **docs/web-app-job-queue-integration.md**: Comprehensive web application integration guide with JavaScript examples
- **docs/SECURE_CREDENTIALS_SETUP.md**: Google Cloud credentials and secrets management setup

### Key Implementation Notes
- **Job Queue Architecture**: See `docs/gcloud-job-queue-implementation.md` for detailed system design and background processing logic
- **Client Integration**: See `docs/web-app-job-queue-integration.md` for polling patterns, error handling, and React hooks
- **Security Setup**: See `docs/SECURE_CREDENTIALS_SETUP.md` for proper secrets management in Google Cloud

### Recent Major Updates
1. **Async Job Queue Implementation**: Complete system for handling long-running operations
2. **Production Configuration**: Optimized for memory-only processing in Cloud Functions
3. **Storage Layer Improvements**: Fixed GCS Content-Type handling and added Local/GCS abstraction
4. **SharePoint Integration**: Centralized file handler for all endpoints
5. **Code Reuse Architecture**: Zero duplication between sync/async processing paths

### Deployment Notes
- **Production**: Uses memory-only processing with Local storage backend
- **API Authentication**: API_KEY stored in Google Secret Manager (excel-pptx-merger-api-key)
- **Secrets Management**: Graph API credentials stored in Google Secret Manager
- **Region**: Deployed to europe-west2 for European users

### Critical Secret Management Gotcha
**IMPORTANT**: When creating secrets with `echo`, ALWAYS use `echo -n` to prevent adding newline characters:
```bash
# CORRECT - prevents newline from being added to secret
echo -n "your-api-key-value" | gcloud secrets create secret-name --data-file=-

# WRONG - adds newline character, causing authentication failures
echo "your-api-key-value" | gcloud secrets create secret-name --data-file=-
```
This issue can cause 401 authentication errors even when the API key appears correct in logs.
