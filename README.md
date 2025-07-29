# Excel to PowerPoint Merger

A robust Python service that extracts data from Excel files and merges it into PowerPoint templates using Jinja-style merge fields. Designed for local development, testing, and Google Cloud Function deployment.

## 🚀 Features

- **Dynamic Excel Data Extraction**: Flexible table detection with configurable search criteria
- **PowerPoint Template Processing**: Jinja-style merge field replacement (`{{field_name}}`)
- **Dynamic Slide Duplication**: Create multiple slides from templates based on list data
- **Slide Filtering**: Include or exclude specific slides in final output
- **Multiple Data Orientations**: Support for horizontal (key-value) and vertical (table) data layouts
- **Image Handling**: Extract images from Excel and replace image placeholders in PowerPoint with aspect ratio preservation
- **Explicit Field Type Support**: Configure field types (text, image, number, date, boolean) for precise data handling
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

### 3. Local Development Folder Structure

The application uses a standardized folder structure for temporary files:

```
.temp/                     # Local development temporary directory
├── excel_pptx_merger_*    # Unique session directory
    ├── input/             # Uploaded Excel and PowerPoint files
    ├── output/            # Generated merged PowerPoint files
    ├── images/            # Extracted images from Excel files
    └── debug/             # Debug information (in development mode)
```

This structure is automatically created when you run the application. The `.temp` directory is included in `.gitignore` to prevent temporary files from being committed.

### 4. Start Development Server

```bash
# Using the convenience script
python scripts/run_local_server.py --debug

# Or using uv directly
uv run python -m src.main serve --debug

# Or using Docker
docker-compose -f docker/docker-compose.yml up
```

### 5. Test the API

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
              "Client": {
                "name": "client_name",
                "type": "text"
              },
              "Word Or Image": {
                "name": "search_type",
                "type": "text"
              }
            }
          }
        },
        {
          "name": "word_search",
          "type": "table",
          "header_search": {
            "method": "contains_text",
            "text": "Word",
            "column": "A",
            "search_range": "A1:A50"
          },
          "data_extraction": {
            "headers_row_offset": 0,
            "data_row_offset": 1,
            "max_rows": 10,
            "column_mappings": {
              "Word": {
                "name": "word",
                "type": "text"
              },
              "Search Criteria": {
                "name": "search_criteria",
                "type": "text"
              }
            }
          }
        }
      ]
    }
  }
}
```

## 📘 User Manual

### PowerPoint Template Design

#### Merge Field Syntax

Merge fields in PowerPoint templates use the Jinja-style double curly braces syntax:

```
{{field_name}}
```

For nested data or arrays, use dot notation:

```
{{table_name.0.field_name}}
```

#### Image Placeholders

To create an image placeholder:

1. Create a text box in PowerPoint
2. Add a merge field with the name of an image field: `{{image_field_name}}`
3. Format the text box to the desired size and position
4. The image will be inserted maintaining its original aspect ratio and centered within the placeholder

#### Best Practices

- **Sizing**: Make image placeholders slightly larger than needed to accommodate various image sizes
- **Text Alternatives**: Consider adding conditional text for cases where images might be missing
- **Field Naming**: Use descriptive field names that match your Excel data structure
- **Testing**: Test templates with sample data to ensure proper field replacement

### Excel Requirements

#### Data Structure

The Excel processor supports two main data structures:

1. **Key-Value Pairs**: For client info and other metadata (horizontal or vertical orientation)
2. **Tables**: For lists of items with multiple columns

#### Image Handling

To include images in your Excel data:

1. Insert images into Excel cells
2. Configure the corresponding fields as `"type": "image"` in the configuration
3. The processor will extract these images and make them available for PowerPoint insertion

#### Best Practices

- **Sheet Names**: Use consistent sheet names that match your configuration
- **Headers**: Include clear headers that match your configuration's search criteria
- **Data Formatting**: Keep data consistent with expected types (text, numbers, dates)
- **Images**: Insert images properly into cells rather than floating them

### Configuration Format

#### Field Type Support

The configuration now supports explicit field type information:

```json
"column_mappings": {
  "Header Name": {
    "name": "field_name",
    "type": "text|image|number|date|boolean"
  }
}
```

Supported field types:
- **text**: Text content (default)
- **image**: Image content (path, URL, or binary data)
- **number**: Numeric values
- **date**: Date values
- **boolean**: Boolean values

#### Configuration Sections

1. **sheet_configs**: Defines which Excel sheets to process
2. **subtables**: Defines tables or key-value sections within each sheet
3. **header_search**: Defines how to locate the start of each data section
4. **data_extraction**: Defines how to extract and map data from the section

#### Example Configurations

See the `config/default_config.json` file for comprehensive examples of:
- Key-value pair extraction
- Table data extraction
- Image field configuration
- Search criteria options

### PowerPoint Features Configuration

#### Dynamic Slide Duplication

Create multiple slides from a template slide based on list data:

```json
{
  "global_settings": {
    "powerpoint": {
      "dynamic_slides": {
        "enabled": true,
        "template_marker": "{{#list:",
        "remove_template_slides": true
      }
    }
  }
}
```

**Template slide example:**
```
{{#list:products}}
Product: {{name}}
Price: {{price}}
Stock: {{quantity}}
```

This will create one slide for each item in the "products" list, replacing merge fields with data from each item.

#### Slide Filtering

Control which slides appear in the final output:

```json
{
  "global_settings": {
    "powerpoint": {
      "slide_filter": {
        "include_slides": [1, 3, 5],  // Only include these slides
        "exclude_slides": [2, 4]      // Or exclude specific slides
      }
    }
  }
}
```

**Notes:**
- Slide numbers are 1-based (matching PowerPoint UI)
- If `include_slides` is specified, only those slides are included
- If only `exclude_slides` is specified, all slides except those are included
- Empty configuration includes all slides

## 🔄 Processing Flow

1. **Excel Processing**:
   - Locate sheets based on configuration
   - Find headers using search criteria
   - Extract data according to extraction rules
   - Apply column mappings and field type information
   - Extract embedded images

2. **PowerPoint Processing**:
   - Load template presentation
   - Find merge fields in text shapes
   - Replace text fields with corresponding data
   - Replace image placeholders with actual images (maintaining aspect ratio)
   - Save the resulting presentation

## 🧪 Testing

```bash
# Run all tests
python -m pytest

# Run specific test file
python -m pytest tests/test_excel_processor.py

# Run with coverage
python -m pytest --cov=src
```

## 🚀 Deployment

### Google Cloud Function

```bash
# Deploy to Google Cloud
python scripts/deploy_gcp.py --project your-project-id
```

### Docker Deployment

```bash
# Build Docker image
docker build -f docker/Dockerfile -t excel-pptx-merger:latest .

# Run Docker container
docker run -p 5000:5000 excel-pptx-merger:latest
```

## 📚 Additional Resources

- **API Documentation**: Available at `/api/v1/` endpoints when server is running
- **Configuration Guide**: See `config/default_config.json` for examples
- **Deployment Guide**: See `scripts/deploy_gcp.py` for Google Cloud deployment

## ⚠️ Security Considerations

- Validate all input files before processing
- Use proper authentication for API endpoints
- Sanitize file paths to prevent directory traversal
- Consider encryption for sensitive data

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 👥 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.