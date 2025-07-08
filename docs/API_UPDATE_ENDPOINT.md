# /api/v1/update Endpoint Documentation

## Overview
The `/api/v1/update` endpoint allows updating Excel files with new data. It supports dual payload modes for maximum compatibility with different client types, including CRM systems and automation tools.

## Endpoint Details
- **URL**: `/api/v1/update`
- **Method**: `POST`
- **Authentication**: API key (if configured)

## Supported Payload Modes

### Mode 1: Multipart Form Data (Standard Clients)

**Use Case**: Postman, web browsers, standard HTTP clients

**Content-Type**: `multipart/form-data`

**Request Structure**:
```
POST /api/v1/update
Content-Type: multipart/form-data; boundary=...

--boundary
Content-Disposition: form-data; name="excel_file"; filename="data.xlsx"
Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet

[BINARY EXCEL FILE DATA]
--boundary
Content-Disposition: form-data; name="update_data"

{
  "report_info": {
    "report_title": "Updated Report"
  },
  "client_info": {
    "client_name": "ABC Corp",
    "search_type": "Word and Image"
  },
  "image_search": [
    {
      "image": "data:image/svg+xml;base64,PHN2ZyB3aWR0aD0i..."
    }
  ]
}
--boundary
Content-Disposition: form-data; name="config"

{
  "options": {
    "preserve_formatting": true
  }
}
--boundary--
```

### Mode 2: JSON Payload (CRM/Automation Systems)

**Use Case**: Deluge CRM, automation systems, webhooks

**Content-Type**: `application/json` or `text/plain` (auto-detected)

**Request Structure**:
```json
POST /api/v1/update
Content-Type: application/json

{
  "excel_file": "UEsDBBQAAAAIAH1r...",
  "filename": "data.xlsx",
  "update_data": {
    "report_info": {
      "report_title": "Updated Report"
    },
    "client_info": {
      "client_name": "ABC Corp",
      "search_type": "Word and Image"
    },
    "image_search": [
      {
        "image": "data:image/svg+xml;base64,PHN2ZyB3aWR0aD0i..."
      }
    ]
  },
  "config": {
    "options": {
      "preserve_formatting": true
    }
  }
}
```

## Request Parameters

### Multipart Mode

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `excel_file` | File | Yes | Binary Excel file (.xlsx) |
| `update_data` | String (JSON) | Yes | JSON string containing update data |
| `config` | String (JSON) | No | JSON string containing configuration options |

### JSON Mode

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `excel_file` | String | Yes | Base64-encoded Excel file content |
| `filename` | String | No | Original filename (for logging and response) |
| `update_data` | Object | Yes | Direct JSON object containing update data |
| `config` | Object | No | Direct JSON object containing configuration options |

## Update Data Structure

```json
{
  "report_info": {
    "report_title": "string"
  },
  "client_info": {
    "client_name": "string",
    "search_type": "string",
    "gs_classes": "string",
    "sic_code": "string",
    "business_nature": "string",
    "countries": "string"
  },
  "search_criteria": [
    {
      "word": "string"
    }
  ],
  "image_search": [
    {
      "image": "data:image/format;base64,..."
    }
  ]
}
```

## Response

### Success Response
**Status Code**: 200

**Content-Type**: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`

**Headers**:
- `Content-Disposition: attachment; filename="updated_filename.xlsx"`

**Body**: Binary Excel file data

### Error Responses

#### 400 Bad Request
```json
{
  "success": false,
  "error": {
    "type": "ValidationError",
    "message": "Excel file is required",
    "code": 400
  }
}
```

#### 413 Request Entity Too Large
```json
{
  "success": false,
  "error": {
    "type": "ValidationError",
    "message": "Request size exceeds maximum allowed size of 100MB (Request size: 150.5MB). Note: CRM systems may inflate request size due to base64 encoding or additional metadata.",
    "code": 413
  }
}
```

#### 500 Internal Server Error
```json
{
  "success": false,
  "error": {
    "type": "ExcelProcessingError",
    "message": "Failed to process Excel file: Invalid cell reference",
    "code": 500
  }
}
```

## CRM System Compatibility

### Automatic Mode Detection
The endpoint automatically detects the payload format:
1. **Standard Detection**: Checks `Content-Type` header and request structure
2. **Fallback Detection**: For CRM systems that send JSON with incorrect headers
3. **Content-Type Tolerance**: Accepts both `application/json` and `text/plain`

### Common CRM Issues Handled
- **Incorrect Content-Type**: Deluge CRM sends JSON with `text/plain` Content-Type
- **Base64 Encoding**: Handles large base64-encoded files and images
- **Payload Size**: Optimized for CRM systems that may add metadata

## Examples

### Curl Example (Multipart)
```bash
curl -X POST "https://your-domain.com/api/v1/update" \
  -H "Authorization: Bearer your-api-key" \
  -F "excel_file=@data.xlsx" \
  -F "update_data={\"client_info\":{\"client_name\":\"ABC Corp\"}}" \
  -F "config={}" \
  --output updated_data.xlsx
```

### Curl Example (JSON)
```bash
curl -X POST "https://your-domain.com/api/v1/update" \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer your-api-key" \
  -d '{
    "excel_file": "UEsDBBQAAAAIAH1r...",
    "filename": "data.xlsx",
    "update_data": {
      "client_info": {
        "client_name": "ABC Corp"
      }
    },
    "config": {}
  }' \
  --output updated_data.xlsx
```

### JavaScript Example (JSON Mode)
```javascript
const response = await fetch('/api/v1/update', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'Authorization': 'Bearer your-api-key'
  },
  body: JSON.stringify({
    excel_file: base64ExcelData,
    filename: 'data.xlsx',
    update_data: {
      client_info: {
        client_name: 'ABC Corp'
      }
    },
    config: {}
  })
});

const blob = await response.blob();
const url = window.URL.createObjectURL(blob);
const a = document.createElement('a');
a.href = url;
a.download = 'updated_data.xlsx';
a.click();
```

## Rate Limiting
- Standard rate limits apply (if configured)
- Large file processing may take longer
- Consider timeout settings for automation systems

## Security
- API key authentication (if enabled)
- File type validation (Excel files only)
- Size limits to prevent abuse
- Input sanitization for all data fields

## Troubleshooting

### Common Issues
1. **413 Error**: Check file size and base64 encoding overhead
2. **Content-Type Issues**: Ensure JSON mode is properly detected
3. **File Corruption**: Verify base64 encoding is correct
4. **Missing Data**: Check JSON structure matches expected format

### Debug Mode
Enable debug mode to get detailed logging:
- Set `DEVELOPMENT_MODE=true`
- Set `LOG_LEVEL=DEBUG`
- Check logs for payload detection and processing details