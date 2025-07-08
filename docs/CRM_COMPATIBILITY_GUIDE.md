# CRM & Automation System Compatibility Guide

## Overview
This guide covers integration with CRM systems and automation platforms that may have non-standard HTTP request behaviors.

## Supported CRM Systems

### ✅ Deluge CRM (Zoho)
- **Status**: Fully supported with automatic detection
- **Issues Resolved**: Content-Type header inconsistencies, base64 payload handling
- **Mode**: JSON with enhanced detection

### ✅ General Webhook Systems
- **Status**: Supported via dual payload mode
- **Compatibility**: Both standard and non-standard implementations

### ✅ Custom Automation Scripts
- **Status**: Flexible payload format support
- **Options**: Multipart or JSON mode based on implementation

## Common CRM Integration Challenges

### 1. Content-Type Header Issues

**Problem**: CRM systems often send JSON data with incorrect Content-Type headers.

**Example**:
```
Content-Type: text/plain
Content-Length: 770416

{"excel_file": "UEsDBBQ...", "update_data": {...}}
```

**Solution**: Enhanced Content-Type detection automatically handles this:
- Attempts standard JSON parsing first
- Falls back to raw data parsing for incorrect headers
- Logs Content-Type mismatches for debugging

### 2. Base64 Encoding Overhead

**Problem**: CRM systems encode binary files as base64, increasing payload size by ~33%.

**Impact**:
- 1MB Excel file → ~1.33MB base64 string
- Additional JSON structure overhead
- Potential size limit issues

**Solution**: 
- Increased form memory limits
- JSON mode bypasses multipart parsing limits
- Automatic base64 detection and decoding

### 3. Payload Structure Variations

**Problem**: Different CRM systems structure payloads differently.

**Examples**:
```json
// Deluge CRM format
{
  "excel_file": "base64...",
  "update_data": {...},
  "config": {...}
}

// Generic webhook format
{
  "file": "base64...",
  "data": {...},
  "options": {...}
}
```

**Solution**: Flexible field mapping and validation with clear error messages.

## Integration Patterns

### Pattern 1: Direct JSON (Recommended for CRM)

**Use Case**: Deluge CRM, Zapier, custom webhooks

**Configuration**:
```javascript
// CRM script example
const payload = {
  excel_file: base64EncodedFile,
  filename: originalFileName,
  update_data: {
    client_info: {...},
    image_search: [{image: "data:image/svg+xml;base64,..."}]
  },
  config: {}
};

fetch('/api/v1/update', {
  method: 'POST',
  headers: {'Content-Type': 'application/json'},
  body: JSON.stringify(payload)
});
```

### Pattern 2: Multipart Form (Standard Clients)

**Use Case**: Postman, web forms, standard HTTP libraries

**Configuration**:
```javascript
const formData = new FormData();
formData.append('excel_file', fileBlob, 'data.xlsx');
formData.append('update_data', JSON.stringify(updateData));
formData.append('config', JSON.stringify(config));

fetch('/api/v1/update', {
  method: 'POST',
  body: formData
});
```

## Deluge CRM Specific Guide

### Setup Instructions

1. **Create Function in Deluge**:
```deluge
// Convert file to base64
fileContent = file.toString("base64");

// Prepare payload
payload = {
  "excel_file": fileContent,
  "filename": file.getName(),
  "update_data": {
    "client_info": {
      "client_name": Client_Name,
      "search_type": Search_Type
    },
    "image_search": imageList
  },
  "config": {}
};

// Make API call
response = invokeurl
[
  url: "https://your-api-domain.com/api/v1/update"
  type: POST
  parameters: payload.toString()
  headers: {"Content-Type": "application/json", "Authorization": "Bearer " + api_key}
];
```

2. **Handle Response**:
```deluge
if(response.get("success") == true) {
  // Process returned Excel file
  updatedFile = response.get("data");
  // Store or attach to record
} else {
  // Handle error
  error = response.get("error");
  info "Error: " + error.get("message");
}
```

### Common Deluge Issues

#### Issue 1: Content-Type Detection
**Symptom**: "Excel file missing from multipart request"
**Cause**: Deluge sends Content-Type as `text/plain`
**Status**: ✅ **Resolved** - Automatic detection implemented

#### Issue 2: Large Base64 Payloads
**Symptom**: 413 Request Entity Too Large
**Cause**: Base64 encoding + multipart overhead
**Status**: ✅ **Resolved** - JSON mode bypasses limits

#### Issue 3: File Corruption
**Symptom**: "Invalid Excel file" errors
**Cause**: Multipart parsing issues with large form fields
**Status**: ✅ **Resolved** - JSON mode processes files correctly

## Troubleshooting Guide

### Debug Mode Setup
Enable detailed logging for CRM integration debugging:

```bash
# Environment variables
DEVELOPMENT_MODE=true
LOG_LEVEL=DEBUG
FLASK_DEBUG=true
```

### Log Analysis

**Normal Operation**:
```
INFO - Update request received - Content-Type: text/plain, Content-Length: 770416
INFO - Detected JSON payload despite Content-Type: text/plain
INFO - Processing request in JSON mode (base64 Excel file)
INFO - Parsed JSON from raw request data due to incorrect Content-Type
INFO - Base64 Excel file size: 1027220 characters
INFO - Decoded Excel file size: 770416 bytes
```

**Error Patterns**:
```
ERROR - Excel file missing from multipart request
→ Solution: Check Content-Type detection

ERROR - Invalid base64 Excel file
→ Solution: Verify base64 encoding in CRM script

ERROR - JSON parsing failed
→ Solution: Check payload structure and quotes
```

### Performance Optimization

#### For Large Files
1. **Compression**: Use gzip compression in CRM requests
2. **Chunking**: Consider file size limits in CRM platform
3. **Timeout**: Increase timeout settings for large files

#### For High Volume
1. **Rate Limiting**: Implement appropriate delays between requests
2. **Batch Processing**: Group multiple updates when possible
3. **Monitoring**: Track API usage and response times

## Security Considerations

### API Key Management
```deluge
// Store API key securely in CRM settings
api_key = zoho.crm.getOrgVariable("EXCEL_API_KEY");

// Use in headers
headers = {
  "Authorization": "Bearer " + api_key,
  "Content-Type": "application/json"
};
```

### Data Validation
- **File Type**: Only Excel files (.xlsx) accepted
- **Size Limits**: Check file size before base64 encoding
- **Content Sanitization**: All text inputs are sanitized

### Error Handling
```deluge
try {
  response = invokeurl[...];
  
  if(response.containsKey("error")) {
    error_message = response.get("error").get("message");
    // Log error and notify user
    info "API Error: " + error_message;
  }
} catch (e) {
  // Handle network/parsing errors
  info "Request failed: " + e.toString();
}
```

## Testing & Validation

### Test Checklist
- [ ] Content-Type detection works with `text/plain`
- [ ] Base64 file encoding/decoding is correct
- [ ] Large payloads (>1MB) process successfully
- [ ] Error messages are clear and actionable
- [ ] Response files are valid and downloadable

### Sample Test Data
```json
{
  "excel_file": "UEsDBBQAAAAIAH1r8VdkNOeVAQAAcAIAABEAAABkb2NQcm9wcy9jb3JlLnhtbJTNPQgCMRTG7ycTI5HFAR1dHBQRHHRUHBr16WOg8/0lLG1+czLu6O9bv...",
  "filename": "test_data.xlsx",
  "update_data": {
    "client_info": {
      "client_name": "Test Client",
      "search_type": "Word and Image"
    },
    "image_search": [
      {
        "image": "data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTUwIiBoZWlnaHQ9IjIwMCIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KICA8cmVjdCB3aWR0aD0iMTUwIiBoZWlnaHQ9IjIwMCIgZmlsbD0iIzMzNzNkYyIvPgogIDx0ZXh0IHg9Ijc1IiB5PSIxMDAiIGZvbnQtZmFtaWx5PSJBcmlhbCIgZm9udC1zaXplPSIxNiIgZmlsbD0id2hpdGUiIHRleHQtYW5jaG9yPSJtaWRkbGUiIGR5PSIwLjNlbSI+TE9HTzE8L3RleHQ+Cjwvc3ZnPg=="
      }
    ]
  },
  "config": {}
}
```

## Support & Maintenance

### Monitoring
- Track CRM request patterns and success rates
- Monitor payload sizes and processing times
- Log Content-Type detection statistics

### Updates
- Test compatibility with CRM platform updates
- Maintain documentation for new integration patterns
- Update error handling based on real-world usage

### Contact
For CRM-specific integration support, include:
- CRM platform and version
- Sample request/response (sanitized)
- Error messages and logs
- Expected vs actual behavior