# Excel Range Images - Usage Guide

This guide explains how to use the Excel Range Images feature to export formatted Excel cell ranges as images and insert them into PowerPoint presentations.

## Overview

The Excel Range Images feature allows you to:
- Export Excel cell ranges as high-quality images
- Preserve formatting, styling, and visual appearance
- Insert range images into PowerPoint templates via merge fields
- Support multiple image formats and DPI settings

## Prerequisites

### Microsoft Graph API Setup

1. **Azure App Registration**: Create an Azure app registration with the following permissions:
   - `Files.ReadWrite.All` (Application permission)
   - `Sites.ReadWrite.All` (Application permission)

2. **Credentials**: Obtain your Azure app credentials:
   - Client ID
   - Client Secret  
   - Tenant ID

### Configuration

Create a Graph API configuration file (copy from template):

```bash
cp config/graph_api_template.env config/graph_api.env
```

Edit `config/graph_api.env` with your credentials:

```env
GRAPH_CLIENT_ID=your_client_id_here
GRAPH_CLIENT_SECRET=your_client_secret_here
GRAPH_TENANT_ID=your_tenant_id_here

# Optional settings
GRAPH_API_TIMEOUT=60
RANGE_EXPORT_DEFAULT_DPI=150
RANGE_EXPORT_DEFAULT_FORMAT=png
```

## Basic Usage

### 1. Configuration Setup

Add range image definitions to your configuration file:

```json
{
  "version": "1.1",
  "sheet_configs": {
    "Data": {
      "subtables": [...]
    }
  },
  "range_images": [
    {
      "field_name": "summary_table",
      "sheet_name": "Data",
      "range": "A1:E15",
      "include_headers": true,
      "output_format": "png",
      "dpi": 150
    }
  ],
  "global_settings": {
    "range_images": {
      "enabled": true,
      "default_dpi": 150,
      "temp_cleanup": true
    }
  }
}
```

### 2. PowerPoint Template Setup

Add range image placeholders in your PowerPoint template:

```
{{summary_table}}
```

The placeholder will be replaced with the exported range image.

### 3. Processing

Use the feature via the API or CLI:

```python
from src.excel_processor import ExcelProcessor
from src.pptx_processor import PowerPointProcessor
from src.graph_api_config import get_graph_api_credentials

# Load Excel file with Graph API credentials
credentials = get_graph_api_credentials()
processor = ExcelProcessor('data.xlsx', graph_credentials=credentials)

# Extract data including range images
data = processor.extract_data(
    global_settings=config['global_settings'],
    sheet_config=config['sheet_configs'],
    full_config=config  # Include range_images
)

# Merge into PowerPoint
pptx_processor = PowerPointProcessor('template.pptx')
pptx_processor.merge_data(data, 'output.pptx')
```

## Configuration Reference

### Range Image Configuration

| Field | Type | Required | Default | Description |
|-------|------|----------|---------|-------------|
| `field_name` | string | Yes | - | Unique identifier for the range image |
| `sheet_name` | string | Yes | - | Excel worksheet name |
| `range` | string | Yes | - | Excel range (e.g., "A1:E15") |
| `include_headers` | boolean | No | true | Whether to include header row |
| `output_format` | string | No | "png" | Image format: "png", "jpg", "jpeg" |
| `dpi` | integer | No | 150 | Image resolution (72-600) |
| `fit_to_content` | boolean | No | true | Auto-fit image to content |
| `width` | integer | No | null | Fixed width in pixels |
| `height` | integer | No | null | Fixed height in pixels |

### Global Settings

```json
{
  "global_settings": {
    "range_images": {
      "enabled": true,
      "max_range_cells": 10000,
      "default_dpi": 150,
      "default_format": "png",
      "temp_cleanup_enabled": true,
      "graph_api": {
        "timeout_seconds": 60,
        "retry_attempts": 3,
        "retry_delay_seconds": 2
      }
    }
  }
}
```

## Advanced Usage

### Multiple Range Images

Export multiple ranges from different sheets:

```json
{
  "range_images": [
    {
      "field_name": "quarterly_summary",
      "sheet_name": "Q4_Data",
      "range": "A1:F20",
      "dpi": 200
    },
    {
      "field_name": "financial_chart",
      "sheet_name": "Charts",
      "range": "B5:L25",
      "output_format": "jpg",
      "width": 800,
      "height": 600
    },
    {
      "field_name": "detailed_breakdown",
      "sheet_name": "Details",
      "range": "A1:K50",
      "dpi": 300,
      "fit_to_content": false
    }
  ]
}
```

### Custom Dimensions

Control image size with custom dimensions:

```json
{
  "field_name": "custom_chart",
  "sheet_name": "Data",
  "range": "A1:H20",
  "fit_to_content": false,
  "width": 1200,
  "height": 800,
  "dpi": 300
}
```

### High-Quality Images

For presentation-quality images:

```json
{
  "field_name": "presentation_chart",
  "sheet_name": "Charts",
  "range": "A1:J30",
  "output_format": "png",
  "dpi": 300,
  "fit_to_content": true
}
```

## PowerPoint Integration

### Placeholder Formats

Range images use standard merge field syntax:

```
{{field_name}}
```

### Multiple Images

Include multiple range images in a single presentation:

```
Quarterly Summary:
{{quarterly_summary}}

Financial Overview:
{{financial_chart}}

Detailed Analysis:
{{detailed_breakdown}}
```

### Image Positioning

The exported image will:
- Replace the placeholder text box
- Maintain aspect ratio (if `fit_to_content: true`)
- Use specified dimensions (if `width`/`height` provided)
- Preserve PowerPoint slide layout

## Error Handling

### Common Issues

1. **Graph API Not Configured**
   ```
   Range exporter not initialized - Graph API credentials required
   ```
   Solution: Set up Graph API credentials in config file

2. **Sheet Not Found**
   ```
   Sheet 'SheetName' not found. Available: ['Sheet1', 'Sheet2']
   ```
   Solution: Check sheet name spelling and case sensitivity

3. **Invalid Range**
   ```
   Range 'A1:InvalidRange' is invalid or empty
   ```
   Solution: Use valid Excel range format (e.g., "A1:C10")

4. **Authentication Errors**
   ```
   Failed to authenticate with Graph API: 401 Unauthorized
   ```
   Solution: Verify Azure app credentials and permissions

### Troubleshooting

Enable debug logging to diagnose issues:

```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

Check the logs for detailed error messages and Graph API responses.

## Performance Considerations

### Range Size Limits

- Maximum recommended: 10,000 cells per range
- Large ranges may take longer to process
- Consider breaking large ranges into smaller chunks

### Optimization Tips

1. **Use appropriate DPI**: Higher DPI = larger files, longer processing
2. **Limit range size**: Smaller ranges process faster
3. **Choose efficient formats**: PNG for quality, JPG for smaller files
4. **Enable temp cleanup**: Prevents disk space issues

### Timeout Settings

Adjust timeouts for large ranges:

```env
GRAPH_API_TIMEOUT=120  # 2 minutes for large ranges
```

## Security Notes

### Credential Management

- Store credentials securely (environment variables or secure config)
- Never commit credentials to version control
- Use least-privilege Azure app permissions
- Rotate credentials regularly

### Data Handling

- Excel files are temporarily uploaded to OneDrive
- Files are automatically cleaned up after processing
- No persistent storage of user data in Microsoft cloud

## Migration Guide

### From Version 1.0

1. Update configuration version to "1.1"
2. Add `range_images` section to configuration
3. Install required dependencies
4. Set up Graph API credentials

### Backward Compatibility

- Existing configurations continue to work
- Range images are an additive feature
- No breaking changes to current API

## API Reference

### ExcelProcessor

```python
# Initialize with Graph API credentials
processor = ExcelProcessor(
    file_input="data.xlsx",
    graph_credentials={
        "client_id": "...",
        "client_secret": "...", 
        "tenant_id": "..."
    }
)

# Extract data with range images
data = processor.extract_data(
    global_settings=global_settings,
    sheet_config=sheet_config,
    full_config=full_config  # Must include range_images
)
```

### Configuration Validation

```python
from src.config_schema_validator import validate_config_file

is_valid, errors = validate_config_file(config)
if not is_valid:
    print("Configuration errors:", errors)
```

## Examples

See the `config/range_images_example_config.json` file for a complete working example.

## Support

For issues and questions:
1. Check the troubleshooting section above
2. Review log files for detailed error messages
3. Verify Graph API credentials and permissions
4. Test with a simple range first