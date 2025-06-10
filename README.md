# Excel to PowerPoint Merger

A robust Python service that extracts data from Excel files and merges it into PowerPoint templates using Jinja-style merge fields. Designed for local development, testing, and Google Cloud Function deployment.

## 🚀 Features

- **Dynamic Excel Data Extraction**: Flexible table detection with configurable search criteria
- **PowerPoint Template Processing**: Jinja-style merge field replacement (`{{field_name}}`)
- **Multiple Data Orientations**: Support for horizontal (key-value) and vertical (table) data layouts
- **Image Handling**: Extract images from Excel and replace image placeholders in PowerPoint
- **Environment-Based Configuration**: Development, testing, and production configurations
- **Temporary File Management**: Intelligent cleanup with configurable retention policies
- **API Endpoints**: RESTful API with health checks, preview, and processing endpoints
- **Google Cloud Ready**: Deploy as Cloud Function with one command
- **Comprehensive Testing**: Unit tests, integration tests, and fixtures
- **Docker Support**: Containerized development and deployment

## 📋 Requirements

- Python 3.9+
- uv package manager (recommended) or pip
- Google Cloud SDK (for deployment)
- Docker (optional, for containerized development)

## 🛠️ Quick Start

### 1. Clone and Setup

```bash
git clone <repository-url>
cd excel_pptx_merger

# Run the automated setup script
python scripts/setup_dev.py
```

### 2. Environment Configuration

```bash
# Copy environment template
cp .env.example .env

# Edit with your configuration
nano .env
```

### 3. Start Development Server

```bash
# Using the convenience script
python scripts/run_local_server.py --debug

# Or using uv directly
uv run python -m src.main serve --debug

# Or using Docker
docker-compose -f docker/docker-compose.yml up
```

### 4. Test the API

```bash
# Health check
curl http://localhost:5000/api/v1/health

# Get default configuration
curl http://localhost:5000/api/v1/config
```

## 🔧 Configuration

### Excel Data Extraction Configuration

The service uses JSON configuration to define how data should be extracted from Excel files:

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

### Environment Variables

```bash
# Application
ENVIRONMENT=development
DEVELOPMENT_MODE=true
LOG_LEVEL=DEBUG

# API
API_KEY=your_api_key_here
MAX_FILE_SIZE_MB=50
ALLOWED_EXTENSIONS=xlsx,pptx

# Temporary Files
CLEANUP_TEMP_FILES=false
TEMP_FILE_RETENTION_SECONDS=3600

# Google Cloud (for deployment)
GOOGLE_CLOUD_PROJECT=your_project_id
GOOGLE_CLOUD_BUCKET=your_bucket_name
```

## 📊 API Reference

### POST /api/v1/merge

Merge Excel data into PowerPoint template.

**Request:**
- `excel_file`: Excel file (.xlsx)
- `pptx_file`: PowerPoint template (.pptx)
- `config`: JSON configuration (optional)

**Response:** Merged PowerPoint file

```bash
curl -X POST http://localhost:5000/api/v1/merge \
  -F "excel_file=@data.xlsx" \
  -F "pptx_file=@template.pptx" \
  -F "config=@config.json" \
  -o merged_output.pptx
```

### POST /api/v1/preview

Preview merge without processing files.

**Response:**
```json
{
  "success": true,
  "preview": {
    "extracted_data": {...},
    "template_info": {...},
    "merge_preview": {...},
    "configuration_used": {...}
  }
}
```

### GET /api/v1/health

Health check endpoint.

**Response:**
```json
{
  "success": true,
  "status": "healthy",
  "version": "0.1.0",
  "timestamp": "2023-12-01T12:00:00Z",
  "services": {
    "config_manager": true,
    "temp_file_manager": true
  }
}
```

### GET /api/v1/config

Get default configuration.

### GET /api/v1/stats

Get system statistics.

## 🎯 Usage Examples

### Basic CLI Usage

```bash
# Merge files using CLI
uv run python -m src.main merge \
  input.xlsx \
  template.pptx \
  output.pptx \
  --config config.json
```

### PowerPoint Template Setup

Create merge fields in your PowerPoint template using double curly braces:

```
Client: {{Order Form.client_info.client_name}}
Type: {{Order Form.client_info.search_type}}
Date: {{Order Form.client_info.order_date}}

Classes:
{{Order Form.trademark_classes.0.class_number}} - {{Order Form.trademark_classes.0.class_description}}
{{Order Form.trademark_classes.1.class_number}} - {{Order Form.trademark_classes.1.class_description}}
```

### Excel Data Structure

Organize your Excel data with clear headers and consistent layouts:

```
Client          Type      Date
Acme Corp       Word      2023-12-01

Class    Description           Status
35       Advertising services  Active
42       Computer services     Pending
```

## 🧪 Testing

```bash
# Run all tests
uv run pytest

# Run with coverage
uv run pytest --cov=src --cov-report=html

# Run specific test file
uv run pytest tests/test_excel_processor.py -v

# Run integration tests only
uv run pytest tests/integration/ -v
```

## 🚀 Deployment

### Google Cloud Functions

```bash
# Set environment variables
export GOOGLE_CLOUD_PROJECT=your-project-id
export API_KEY=your-secure-api-key

# Deploy to Google Cloud
python scripts/deploy_gcp.py \
  --project-id your-project-id \
  --region us-central1 \
  --function-name excel-pptx-merger \
  --env production
```

### Docker Deployment

```bash
# Build and run with Docker
docker-compose -f docker/docker-compose.yml up --build

# Production deployment
docker build -f docker/Dockerfile -t excel-pptx-merger .
docker run -p 8080:8080 \
  -e ENVIRONMENT=production \
  -e API_KEY=your-api-key \
  excel-pptx-merger
```

## 🏗️ Project Structure

```
excel-pptx-merger/
├── src/                          # Source code
│   ├── main.py                   # Main API handler & CLI
│   ├── excel_processor.py        # Excel data extraction
│   ├── pptx_processor.py         # PowerPoint processing
│   ├── config_manager.py         # Configuration management
│   ├── temp_file_manager.py      # Temporary file handling
│   └── utils/                    # Utility modules
│       ├── exceptions.py         # Custom exceptions
│       ├── file_utils.py         # File handling utilities
│       └── validation.py         # Input validation
├── tests/                        # Test suite
│   ├── fixtures/                 # Test data and configurations
│   ├── integration/              # Integration tests
│   ├── test_excel_processor.py   # Excel processor tests
│   └── test_config_manager.py    # Configuration tests
├── config/                       # Configuration files
│   ├── default_config.json       # Default extraction config
│   ├── development.env           # Development environment
│   ├── testing.env              # Testing environment
│   └── production.env           # Production environment
├── scripts/                      # Deployment and utility scripts
│   ├── setup_dev.py             # Development setup
│   ├── run_local_server.py      # Local server runner
│   └── deploy_gcp.py            # Google Cloud deployment
├── docker/                       # Docker configuration
│   ├── Dockerfile               # Docker image definition
│   └── docker-compose.yml       # Local development stack
├── docs/                         # Documentation
│   └── excel_pptx_merger_prd.md # Product Requirements Document
├── pyproject.toml               # Project configuration
├── requirements.txt             # Dependencies for GCP
└── README.md                    # This file
```

## 🔧 Development

### Code Quality

```bash
# Format code
uv run black src/ tests/

# Check style
uv run flake8 src/ tests/

# Type checking
uv run mypy src/

# Run all quality checks
pre-commit run --all-files
```

### Adding New Features

1. **Excel Processing**: Extend `ExcelProcessor` class in `src/excel_processor.py`
2. **PowerPoint Processing**: Extend `PowerPointProcessor` class in `src/pptx_processor.py`
3. **API Endpoints**: Add new routes in `src/main.py`
4. **Configuration**: Update schema in `src/utils/validation.py`
5. **Tests**: Add tests in appropriate `tests/` subdirectory

### Configuration Schema

The configuration system supports:

- **Multiple sheets** with different extraction rules
- **Flexible header search** (contains_text, exact_match, regex)
- **Data orientation** (horizontal key-value pairs, vertical tables)
- **Column mapping** for consistent output keys
- **Image extraction** and replacement
- **Environment-specific overrides**

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Run tests and quality checks
4. Commit changes (`git commit -m 'Add amazing feature'`)
5. Push to branch (`git push origin feature/amazing-feature`)
6. Open a Pull Request

## 📚 Documentation

- **Product Requirements**: [docs/excel_pptx_merger_prd.md](docs/excel_pptx_merger_prd.md)
- **API Documentation**: Available at `/api/v1/` endpoints when server is running
- **Configuration Guide**: See `config/default_config.json` for examples
- **Deployment Guide**: See `scripts/deploy_gcp.py` for Google Cloud deployment

## ⚠️ Security Considerations

- **API Key Authentication**: Set `API_KEY` environment variable for production
- **File Size Limits**: Configure `MAX_FILE_SIZE_MB` to prevent abuse
- **Temporary File Cleanup**: Ensure `CLEANUP_TEMP_FILES=true` in production
- **Input Validation**: All inputs are validated before processing
- **Error Handling**: Sensitive information is not exposed in error messages

## 🐛 Troubleshooting

### Common Issues

**Excel file not found**: Verify file path and permissions
```bash
ls -la path/to/file.xlsx
```

**PowerPoint template invalid**: Check template format and merge field syntax
```bash
# Test with minimal template containing {{test}} field
```

**Memory issues with large files**: Increase container memory or implement streaming
```yaml
# docker-compose.yml
deploy:
  resources:
    limits:
      memory: 2G
```

**Google Cloud deployment fails**: Check authentication and project settings
```bash
gcloud auth list
gcloud config get-value project
```

### Debug Mode

Enable debug mode for detailed logging:

```bash
export DEVELOPMENT_MODE=true
export LOG_LEVEL=DEBUG
python scripts/run_local_server.py --debug
```

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

- **OpenPyXL**: Excel file processing
- **python-pptx**: PowerPoint file manipulation
- **Flask**: Web framework
- **Google Cloud Functions**: Serverless deployment platform
- **uv**: Fast Python package manager

---

**Made with ❤️ for Trademark Helpline**

For support, please contact the development team or create an issue in the repository.