# Excel to PowerPoint Merger - Development Setup Prompt for Claude Code

## Project Overview
Create a complete Python project for the Excel to PowerPoint Merger service that extracts data from Excel files and merges it into PowerPoint templates using Jinja-style merge fields. The project uses uv package manager and should be ready for local development/testing with future Google Cloud Function deployment.

## Project Structure Required
```
excel-pptx-merger/
├── docs/
│   └── Excel to PowerPoint Merger - Product Requirements Document.md  # Already exists
├── src/
│   ├── __init__.py
│   ├── main.py                    # Main Cloud Function entry point
│   ├── excel_processor.py         # Excel data extraction logic
│   ├── pptx_processor.py          # PowerPoint template processing
│   ├── config_manager.py          # Configuration handling
│   ├── temp_file_manager.py       # Temporary file management
│   └── utils/
│       ├── __init__.py
│       ├── file_utils.py          # File handling utilities
│       ├── validation.py          # Input validation
│       └── exceptions.py          # Custom exceptions
├── tests/
│   ├── __init__.py
│   ├── test_excel_processor.py
│   ├── test_pptx_processor.py
│   ├── test_config_manager.py
│   ├── test_temp_file_manager.py
│   ├── fixtures/
│   │   ├── sample_excel.xlsx      # Test Excel file
│   │   ├── sample_template.pptx   # Test PowerPoint template
│   │   └── test_config.json       # Test configuration
│   └── integration/
│       ├── __init__.py
│       └── test_end_to_end.py
├── config/
│   ├── default_config.json        # Default data extraction configuration
│   ├── development.env            # Development environment variables
│   ├── testing.env               # Testing environment variables
│   └── production.env            # Production environment variables
├── scripts/
│   ├── setup_dev.py              # Development environment setup
│   ├── run_local_server.py       # Local Flask server for testing
│   └── deploy_gcp.py             # Google Cloud Function deployment script
├── .env.example                  # Environment variables template
├── .gitignore                    # Git ignore file
├── pyproject.toml               # uv package configuration
├── requirements.txt             # Dependencies for GCP deployment
├── README.md                    # Project documentation
└── docker/
    ├── Dockerfile               # For local containerized testing
    └── docker-compose.yml       # Local development environment
```

## Technical Requirements

### Dependencies to Include
- **Core Libraries**: pandas, openpyxl, python-pptx, Pillow
- **Web Framework**: Flask (for local testing) or functions-framework (for GCP)
- **HTTP Client**: requests (for Zoho WorkDrive integration)
- **Utilities**: python-dotenv, jsonschema, click
- **Development**: pytest, pytest-cov, black, flake8, mypy
- **Cloud**: google-cloud-storage, google-cloud-logging

### Key Features to Implement

#### 1. Excel Data Extraction (excel_processor.py)
- Dynamic table detection with configurable search criteria
- Support for multiple subtables on single sheet
- Column name normalization (spaces → underscores, lowercase)
- Image extraction from embedded Excel images
- Flexible end condition detection (empty rows, text markers, max rows)
- Support for horizontal (key-value) and vertical (table) orientations

#### 2. PowerPoint Processing (pptx_processor.py)
- Jinja-style merge field replacement: `{{field_name}}`
- Text placeholder replacement with data
- Image placeholder replacement (text boxes → actual images)
- Maintain original formatting and positioning
- Support for nested object references: `{{table.0.field_name}}`

#### 3. Configuration Management (config_manager.py)
- JSON schema validation for extraction configurations
- Default configuration loading
- Custom configuration override support
- Environment-specific settings

#### 4. Temporary File Management (temp_file_manager.py)
- Environment-based cleanup control
- Development mode (keep files for debugging)
- Production mode (automatic cleanup)
- Configurable retention periods
- Error handling (keep files on processing failures)
- Runtime parameter override support

#### 5. Main API Handler (main.py)
- Flask app for local development
- Cloud Function handler for GCP deployment
- File upload handling (Excel + PowerPoint)
- JSON configuration parameter support
- Comprehensive error handling and logging
- Health check endpoint

### Configuration Schema to Support
```json
{
  "version": "1.0",
  "sheet_configs": {
    "Order Form": {
      "subtables": [
        {
          "name": "client_info",
          "type": "key_value_pairs",
          "header_search": {
            "method": "contains_text",
            "text": "Client",
            "column": "A",
            "search_range": "A1:A10"
          },
          "data_extraction": {
            "orientation": "horizontal",
            "headers_row_offset": 0,
            "data_row_offset": 1,
            "max_columns": 6,
            "column_mappings": {
              "Client": "client_name",
              "Word Or Image": "search_type"
            }
          }
        }
      ]
    }
  },
  "global_settings": {
    "normalize_keys": true,
    "temp_file_cleanup": {
      "enabled": true,
      "delay_seconds": 300,
      "keep_on_error": true,
      "development_mode": false
    }
  }
}
```

### Environment Variables to Support
```bash
# Development
DEVELOPMENT_MODE=true
CLEANUP_TEMP_FILES=false
TEMP_FILE_RETENTION_SECONDS=3600
LOG_LEVEL=DEBUG

# Production
DEVELOPMENT_MODE=false
CLEANUP_TEMP_FILES=true
TEMP_FILE_RETENTION_SECONDS=60
LOG_LEVEL=INFO

# API Configuration
API_KEY=your_api_key_here
MAX_FILE_SIZE_MB=50
ALLOWED_EXTENSIONS=xlsx,pptx

# Zoho WorkDrive (for future integration)
ZOHO_CLIENT_ID=your_client_id
ZOHO_CLIENT_SECRET=your_client_secret
ZOHO_REFRESH_TOKEN=your_refresh_token
```

### API Endpoints to Implement
1. **POST /api/v1/merge** - Main file processing endpoint
2. **GET /api/v1/health** - Health check and status
3. **POST /api/v1/config** - Store configuration (future)
4. **GET /api/v1/config/{name}** - Retrieve configuration (future)

### Testing Requirements
- Unit tests for all major components
- Integration tests for end-to-end workflows
- Test fixtures with sample Excel and PowerPoint files
- Mock external dependencies (Zoho API)
- Performance testing for large files
- Error scenario testing

### Local Development Features
- Flask development server with hot reload
- Detailed logging and debug output
- Temporary file inspection capabilities
- Configuration validation tools
- Sample file generation utilities

### Deployment Preparation
- Google Cloud Function compatible structure
- requirements.txt for GCP deployment
- Environment variable management
- Deployment scripts and documentation

## Implementation Notes

### Column Name Normalization Logic
```python
def normalize_column_name(column_name):
    """Convert Excel headers to JSON keys"""
    # Convert to lowercase, replace spaces with underscores
    # Remove special characters, handle edge cases
    # Examples: "Client Name" → "client_name", "G&S Classes" → "g_s_classes"
```

### Error Handling Strategy
- Comprehensive exception handling with detailed error messages
- Graceful fallbacks for missing data or configuration issues
- Proper HTTP status codes for API responses
- Detailed logging for debugging and monitoring

### File Processing Flow
1. Validate uploaded files (format, size)
2. Extract configuration or use defaults
3. Create temporary working directory
4. Extract Excel data using dynamic table detection
5. Process images and create normalized JSON structure
6. Load PowerPoint template and replace merge fields
7. Generate final presentation
8. Handle temporary file cleanup based on configuration
9. Return processed file or error response

## Deliverables Expected
1. Complete project structure with all files
2. Working local development environment
3. Comprehensive test suite with sample data
4. Documentation and setup instructions
5. Docker configuration for containerized testing
6. Deployment scripts for Google Cloud Functions
7. Example usage and API documentation

Please create all files with proper error handling, logging, type hints, and documentation. The code should be production-ready but with extensive development and debugging capabilities.