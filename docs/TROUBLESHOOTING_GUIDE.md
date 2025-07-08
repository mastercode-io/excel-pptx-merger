# Troubleshooting Guide

## Common Issues and Solutions

### 413 Request Entity Too Large

#### Symptoms
- Error: "413 Request Entity Too Large: The data value transmitted exceeds the capacity limit"
- CRM requests fail while Postman works
- Works with small files but fails with larger payloads

#### Root Causes

##### 1. **Multipart Form Field Limits** ✅ RESOLVED
**Issue**: Werkzeug's default form memory limit (500KB) was too small for base64-encoded images.

**Previous Error Pattern**:
```
Content-Type: multipart/form-data
Form field 'update_data' size: 690158 bytes  # > 500KB limit
```

**Solution Implemented**:
- Increased `MAX_FORM_MEMORY_SIZE` to 10MB
- Added JSON mode to bypass multipart parsing entirely

##### 2. **Content-Type Detection Issues** ✅ RESOLVED
**Issue**: CRM systems sending JSON with `Content-Type: text/plain`.

**Previous Error Pattern**:
```
Content-Type: text/plain, Content-Length: 770416
Request analysis - JSON: False, Form data: False, Files: False
Processing request in multipart mode (binary Excel file)
Excel file missing from multipart request
```

**Solution Implemented**:
- Enhanced JSON detection regardless of Content-Type header
- Automatic fallback to JSON parsing for CRM systems
- Robust Content-Type handling

#### Current Status: ✅ **RESOLVED**
The 413 error has been resolved through:
1. **Dual payload mode** supporting both multipart and JSON
2. **Enhanced Content-Type detection** for CRM compatibility
3. **Increased memory limits** for large form fields
4. **Automatic mode detection** based on request structure

#### Troubleshooting Steps

If you encounter 413 errors:

1. **Check Content-Type**:
```bash
# Look for this in logs
Content-Type: text/plain  # Should trigger JSON mode
```

2. **Verify JSON Mode Detection**:
```bash
# Should see these logs
Detected JSON payload despite Content-Type: text/plain
Processing request in JSON mode (base64 Excel file)
```

3. **Monitor Payload Size**:
```bash
# Check actual vs configured limits
Actual request content length: 770416 bytes (0.73 MB)
Flask configuration - MAX_CONTENT_LENGTH: 100.0MB
```

### File Processing Errors

#### Excel File Missing/Corrupted

**Symptoms**:
- "Excel file missing from request"
- "Invalid Excel file format"
- File exists but appears corrupted

**Root Cause**: Multipart parsing issues with large payloads.

**Solution**: Use JSON mode which bypasses multipart parsing:
```json
{
  "excel_file": "UEsDBBQ...",  // Base64 encoded
  "filename": "data.xlsx",
  "update_data": {...}
}
```

#### Base64 Encoding Issues

**Symptoms**:
- "Invalid base64 Excel file"
- File size discrepancies
- Decoding errors

**Debugging**:
```bash
# Check these log entries
Base64 Excel file size: 1027220 characters
Decoded Excel file size: 770416 bytes
```

**Common Fixes**:
1. **Verify Base64 Encoding**:
```javascript
// Correct encoding
const base64 = btoa(String.fromCharCode(...new Uint8Array(fileBuffer)));

// Or in Deluge CRM
fileContent = file.toString("base64");
```

2. **Check Data Integrity**:
```bash
# Verify base64 string doesn't contain newlines or spaces
echo "UEsDBBQ..." | base64 -d > test.xlsx
```

### CRM Integration Issues

#### Deluge CRM Specific

**Issue**: Content-Type and payload structure incompatibilities.

**Status**: ✅ **RESOLVED** - Automatic detection implemented.

**Verification Logs**:
```bash
Update request User-Agent: Deluge
Detected JSON payload despite Content-Type: text/plain
JSON payload detected with non-standard Content-Type: text/plain
Processing request in JSON mode (base64 Excel file)
```

#### Generic Webhook Issues

**Common Problems**:
1. **Timeout Errors**: Increase timeout in webhook configuration
2. **Authentication**: Ensure API key is correctly set
3. **Response Handling**: Check for binary response handling

**Example Fix (JavaScript)**:
```javascript
const response = await fetch('/api/v1/update', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'Authorization': 'Bearer your-api-key'
  },
  body: JSON.stringify(payload)
});

// Handle binary response
const blob = await response.blob();
```

### API Response Issues

#### Empty or Corrupted Response Files

**Symptoms**:
- Downloaded file is empty
- File opens but contains no data
- Corrupted Excel file

**Debugging Steps**:
1. **Check Response Headers**:
```bash
Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
Content-Disposition: attachment; filename="updated_data.xlsx"
```

2. **Verify Processing Logs**:
```bash
Excel update processing completed successfully
Saved base64-decoded Excel file to: /tmp/input_abc123.xlsx
```

3. **Test with Minimal Data**:
```json
{
  "excel_file": "UEsDBBQ...",
  "update_data": {
    "client_info": {
      "client_name": "Test"
    }
  }
}
```

#### Authentication Errors

**Symptoms**:
- 401 Unauthorized
- "Invalid API key"

**Solutions**:
1. **Check API Key Format**:
```bash
Authorization: Bearer your-api-key-here
```

2. **Verify Environment Variables**:
```bash
API_KEY=your-actual-key
DEVELOPMENT_MODE=false  # Disables auth bypass
```

3. **Test Authentication**:
```bash
curl -H "Authorization: Bearer your-key" /api/v1/health
```

### Performance Issues

#### Slow Processing

**Symptoms**:
- Requests timeout
- Long response times
- Memory issues

**Optimization Steps**:

1. **Check File Sizes**:
```bash
# Monitor these logs
Base64 Excel file size: 1027220 characters  # Large files take longer
Decoded Excel file size: 770416 bytes
```

2. **Optimize Payload**:
- Compress Excel files before base64 encoding
- Remove unnecessary images from payloads
- Use minimal update data structures

3. **Monitor Memory Usage**:
```bash
# Check for memory warnings in logs
Processing Excel update request for file: large_data.xlsx
Mode: JSON (base64)
```

#### Timeout Issues

**Configuration**:
```yaml
# cloudbuild.yaml
--timeout=540s  # 9 minutes for Google Cloud Functions
```

**Client-Side Timeouts**:
```javascript
// Increase timeout for large files
const controller = new AbortController();
setTimeout(() => controller.abort(), 600000); // 10 minutes

fetch('/api/v1/update', {
  signal: controller.signal,
  // ... other options
});
```

## Debug Mode Setup

### Enable Enhanced Logging

**Environment Variables**:
```bash
DEVELOPMENT_MODE=true
LOG_LEVEL=DEBUG
FLASK_DEBUG=true
```

**What You'll See**:
```bash
# Request analysis
Content-Type: text/plain, Content-Length: 770416
Detected JSON payload despite Content-Type: text/plain
Processing request in JSON mode (base64 Excel file)

# File processing
Base64 Excel file size: 1027220 characters
Decoded Excel file size: 770416 bytes
Saved base64-decoded Excel file to: /tmp/input_abc123.xlsx

# Success confirmation
Excel update processing completed successfully
```

### Log Analysis

#### Normal Operation Pattern
```bash
INFO - Update request received - Content-Type: application/json
INFO - Request analysis - JSON: True, Form data: False, Files: False
INFO - Processing request in JSON mode (base64 Excel file)
INFO - Base64 Excel file size: 1027220 characters
INFO - Decoded Excel file size: 770416 bytes
INFO - Excel update processing completed successfully
```

#### Error Patterns

**Content-Type Issues**:
```bash
INFO - Content-Type: text/plain  # CRM system
INFO - Request analysis - JSON: False  # Initially not detected
INFO - Detected JSON payload despite Content-Type: text/plain  # Fixed
INFO - Processing request in JSON mode (base64 Excel file)  # Success
```

**File Processing Errors**:
```bash
ERROR - Invalid base64 Excel file: Invalid base64-encoded string
ERROR - Excel update failed: No such file or directory
ERROR - JSON parsing failed: Expecting ',' delimiter
```

## Monitoring and Alerts

### Key Metrics to Track

1. **Request Success Rate**:
   - JSON mode vs Multipart mode success rates
   - CRM vs standard client success rates

2. **Payload Sizes**:
   - Average base64 file sizes
   - Processing times by file size

3. **Error Patterns**:
   - 413 error frequency (should be 0 after fixes)
   - Content-Type detection accuracy
   - File corruption incidents

### Performance Baselines

**Normal Operation**:
- Files < 1MB: < 10 seconds processing
- Files 1-10MB: < 30 seconds processing
- Success rate: > 99% for properly formatted requests

**Alert Thresholds**:
- 413 errors: Any occurrence (should be resolved)
- Processing time: > 60 seconds for < 10MB files
- Error rate: > 5% over 10-minute window

## Getting Help

### Information to Include

When reporting issues, include:

1. **Request Details**:
   - Client type (Postman, Deluge CRM, etc.)
   - Content-Type header
   - Approximate payload size

2. **Error Messages**:
   - Complete error response
   - Relevant log entries
   - Timestamp of the request

3. **Environment**:
   - Deployment type (local, Google Cloud, etc.)
   - Configuration settings
   - Debug mode enabled/disabled

### Sample Support Request

```
Subject: 413 Error with Deluge CRM Integration

Environment: Google Cloud Function, Production
Client: Deluge CRM
Issue: 413 Request Entity Too Large

Request Details:
- Content-Type: text/plain
- Payload size: ~770KB
- User-Agent: Deluge

Error Message:
"413 Request Entity Too Large: The data value transmitted exceeds the capacity limit"

Logs:
2025-07-08 19:02:21 - Content-Type: text/plain, Content-Length: 770416
2025-07-08 19:02:21 - Request analysis - JSON: False, Form data: False, Files: False

Expected: Request should be processed in JSON mode
Actual: Falls back to multipart mode and fails
```

This format helps identify the issue quickly and provide targeted solutions.