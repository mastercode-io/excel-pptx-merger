# Merge Endpoint Dual Mode Support

## Overview
The `/api/v1/merge` endpoint now supports dual payload modes, matching the flexibility of the `/update` endpoint. This allows CRM systems and automation tools to send base64-encoded files via JSON payloads.

## Supported Modes

### 1. Multipart Mode (Standard)
- **Content-Type**: `multipart/form-data`
- **Files**: Binary file uploads
- **Configuration**: JSON string in form field

Example:
```bash
curl -X POST http://localhost:5000/api/v1/merge \
  -F "excel_file=@sample.xlsx" \
  -F "pptx_file=@template.pptx" \
  -F 'config={"global_settings":{"image_extraction":{"enabled":true}}}'
```

### 2. JSON Mode (CRM/Automation)
- **Content-Type**: `application/json` or `text/plain`
- **Files**: Base64-encoded strings
- **Configuration**: Direct JSON object

Example payload:
```json
{
  "excel_file": "base64_encoded_excel_content",
  "pptx_file": "base64_encoded_pptx_content",
  "excel_filename": "data.xlsx",
  "pptx_filename": "template.pptx",
  "config": {
    "global_settings": {
      "image_extraction": {
        "enabled": true
      }
    }
  }
}
```

## Key Features

1. **Automatic Detection**: The endpoint automatically detects the payload type
2. **CRM Compatibility**: Handles incorrect Content-Type headers (e.g., `text/plain` with JSON data)
3. **Enhanced Logging**: Detailed request analysis for debugging
4. **Unified Processing**: Both modes converge to the same processing pipeline
5. **Memory/File Modes**: Works with both `save_files=true` and memory-only processing

## Testing

Run the test script to verify both modes:
```bash
python test_merge_json_mode.py
```

This will test:
- JSON mode with correct Content-Type
- JSON mode with text/plain Content-Type (CRM compatibility)
- Traditional multipart mode

## Implementation Details

The implementation adds:
- Enhanced request detection logic
- JSON payload parsing with base64 decoding
- Unified file handling for both BytesIO and uploaded files
- Comprehensive error handling for malformed requests
- Detailed logging for troubleshooting