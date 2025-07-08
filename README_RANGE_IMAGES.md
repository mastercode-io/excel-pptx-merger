# Excel Range Images Feature

## Quick Start

The Excel Range Images feature allows you to export formatted Excel ranges as images and insert them into PowerPoint presentations, preserving all visual formatting and styling.

### 1. Setup Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com) → App Registrations
2. Create new registration with these permissions:
   - `Files.ReadWrite.All` (Application)
   - `Sites.ReadWrite.All` (Application)
3. Copy Client ID, Client Secret, and Tenant ID

### 2. Configure Credentials

```bash
cp config/graph_api_template.env config/graph_api.env
# Edit config/graph_api.env with your credentials
```

### 3. Update Configuration

Add range images to your config:

```json
{
  "version": "1.1",
  "range_images": [
    {
      "field_name": "summary_table",
      "sheet_name": "Data", 
      "range": "A1:E15",
      "dpi": 150
    }
  ],
  "global_settings": {
    "range_images": {
      "enabled": true
    }
  }
}
```

### 4. Add PowerPoint Placeholder

In your PowerPoint template:
```
{{summary_table}}
```

### 5. Process

```python
from src.excel_processor import ExcelProcessor
from src.graph_api_config import get_graph_api_credentials

credentials = get_graph_api_credentials()
processor = ExcelProcessor('data.xlsx', graph_credentials=credentials)
data = processor.extract_data(global_settings, sheet_config, full_config)
```

## Features

✅ **High-Quality Image Export** - Preserves Excel formatting, colors, and styling  
✅ **Multiple Format Support** - PNG, JPG with customizable DPI (72-600)  
✅ **Flexible Sizing** - Auto-fit or custom dimensions  
✅ **Batch Processing** - Export multiple ranges in one operation  
✅ **Error Handling** - Robust retry logic and comprehensive error reporting  
✅ **Temporary File Management** - Automatic cleanup of temp files  
✅ **Cross-Platform** - Works on any platform with internet access  

## Architecture

```
Excel File → Graph API → OneDrive → Range Rendering → Image Export → PowerPoint
```

The system:
1. Uploads Excel file to OneDrive temporarily
2. Uses Microsoft Graph API to render ranges as images
3. Downloads high-quality images
4. Inserts images into PowerPoint templates
5. Cleans up temporary files automatically

## Configuration Options

### Range Image Settings
- `field_name`: Unique identifier for PowerPoint placeholder
- `sheet_name`: Excel worksheet name
- `range`: Excel range notation (e.g., "A1:E15")
- `dpi`: Image resolution (72-600, default: 150)
- `output_format`: "png", "jpg", or "jpeg"
- `width`/`height`: Custom dimensions in pixels
- `include_headers`: Include header row in export

### Global Settings
- `enabled`: Enable/disable range image feature
- `max_range_cells`: Maximum cells per range (default: 10,000)
- `default_dpi`: Default image resolution
- `timeout_seconds`: Graph API timeout
- `retry_attempts`: Number of retry attempts

## Use Cases

### Financial Reports
Export formatted financial tables with styling preserved:
```json
{
  "field_name": "quarterly_financials",
  "sheet_name": "Q4_Results",
  "range": "A1:H25",
  "dpi": 300
}
```

### Charts and Visualizations
Include Excel charts in presentations:
```json
{
  "field_name": "sales_chart", 
  "sheet_name": "Charts",
  "range": "B5:M20",
  "width": 800,
  "height": 600
}
```

### Data Tables
Export complex data tables with formatting:
```json
{
  "field_name": "performance_metrics",
  "sheet_name": "Metrics",
  "range": "A1:K50",
  "dpi": 200,
  "output_format": "png"
}
```

## Comparison with Alternatives

| Feature | Range Images | xlwings | openpyxl |
|---------|-------------|---------|----------|
| Cross-platform | ✅ | ❌ (Windows/Mac only) | ✅ |
| Excel installation required | ❌ | ✅ | ❌ |
| Preserves formatting | ✅ | ✅ | ❌ |
| Cloud-based processing | ✅ | ❌ | ❌ |
| High-quality images | ✅ | ✅ | ❌ |
| Batch processing | ✅ | ❌ | ❌ |

## Security & Privacy

- **Temporary Upload**: Files uploaded to OneDrive temporarily during processing
- **Automatic Cleanup**: Files automatically deleted after processing
- **Secure Authentication**: Uses OAuth 2.0 client credentials flow
- **No Persistent Storage**: No long-term storage of user data
- **Audit Trail**: Comprehensive logging of all operations

## Performance

- **Small ranges** (< 100 cells): ~2-5 seconds
- **Medium ranges** (100-1000 cells): ~5-15 seconds  
- **Large ranges** (1000-10000 cells): ~15-60 seconds
- **Concurrent processing**: Supports multiple ranges in parallel

## Dependencies

Required packages:
- `requests` - HTTP client for Graph API
- `Pillow` - Image processing
- `python-pptx` - PowerPoint processing
- `openpyxl` - Excel reading (existing)

## Documentation

- [Complete Usage Guide](docs/EXCEL_RANGE_IMAGES_USAGE.md)
- [Feature Specification](docs/FEATURE_EXCEL_RANGE_IMAGES.md)
- [API Reference](docs/)
- [Configuration Examples](config/range_images_example_config.json)

## Testing

Run the test suite:
```bash
python -m pytest tests/test_graph_api_client.py
python -m pytest tests/test_excel_range_exporter.py
python -m pytest tests/test_config_schema_validator.py
```

## Support

For setup assistance or troubleshooting:
1. Check [Usage Guide](docs/EXCEL_RANGE_IMAGES_USAGE.md)
2. Review log output for detailed error messages
3. Verify Azure app permissions and credentials
4. Test with simple configuration first

---

*This feature requires an active Microsoft 365 subscription and Azure app registration with appropriate permissions.*