# Investigation and Fix Plan for 413 Error with Deluge CRM

## Problem Overview
The `/update` endpoint works correctly in Postman but fails with 413 "Request Entity Too Large" error when called from Deluge CRM automation, despite payload being under 1MB.

## Root Cause Analysis

### File Size Limits Configuration
- **Development**: 50MB limit (`MAX_FILE_SIZE_MB=50`)
- **Production**: 100MB limit (`MAX_FILE_SIZE_MB=100`)  
- **Flask configuration**: `app.config["MAX_CONTENT_LENGTH"] = app_config["max_file_size_mb"] * 1024 * 1024`

### 413 Error Handling Flow
The error originates from Flask's built-in request size validation:
1. Request exceeds `MAX_CONTENT_LENGTH` 
2. Werkzeug raises `RequestEntityTooLarge` exception
3. Custom error handler in `src/main.py:110-118` catches it
4. Returns standardized JSON error response

### Key Differences: Postman vs Deluge

**Postman (Works):**
- Sends standard `multipart/form-data` requests
- Proper Content-Length headers
- Direct file uploads with form fields

**Deluge CRM (Fails):**
- Known for aggressive request transformations
- May add additional headers, encoding, or metadata
- Could be sending requests as `application/json` with base64-encoded files
- May include extra CRM-specific payload data

### Critical Discovery: /update Endpoint Payload Handling

The `/update` endpoint expects:
```python
# Form data extraction (lines 1043-1052)
update_data_str = request.form.get("update_data", "{}")
config_str = request.form.get("config", "{}")

# JSON parsing
update_data = json.loads(update_data_str)
config = json.loads(config_str)
```

**Potential Issues:**
1. **JSON-in-Form vs Pure JSON**: Deluge might send pure JSON payload instead of form-encoded JSON strings
2. **Base64 File Encoding**: CRM systems often base64-encode files, inflating size by ~33%
3. **Additional Metadata**: Deluge may include tracking/audit data in the payload

## Phase 1: Immediate Diagnostic Enhancement

### 1. Add Detailed Request Logging
- Content-Type headers from Deluge requests
- Request payload size breakdown
- Form vs JSON payload detection
- Raw request size vs processed size

### 2. Enhance Error Handling for 413 Errors
- Log actual request size when 413 occurs
- Differentiate between file size vs total payload size
- Add CRM-specific error messages

## Phase 2: Payload Handling Improvements

### 1. Make /update Endpoint More Flexible
- Support both form-encoded and pure JSON payloads
- Auto-detect request format (form vs JSON)
- Handle both direct file uploads and base64-encoded files

### 2. Add Request Size Analysis
- Calculate actual vs inflated payload sizes
- Warn when base64 encoding is detected
- Provide size optimization suggestions

## Phase 3: Configuration Adjustments

### 1. Add Environment-Specific Overrides
- Separate limits for different client types
- CRM-specific payload handling modes
- Flexible content-type handling

### 2. Implement Request Preprocessing
- Detect and handle CRM-specific transformations
- Optimize payload size before processing
- Add fallback mechanisms for oversized requests

## Phase 4: Testing and Validation

### 1. Create CRM Simulation Tests
- Test base64-encoded file uploads
- Simulate Deluge payload transformations
- Validate error handling and logging

### 2. Monitor and Adjust
- Track request patterns from different clients
- Fine-tune size limits based on real usage
- Implement client-specific optimizations

## Expected Outcomes
This plan will resolve the 413 error by making the endpoint more robust to different payload formats while maintaining security and performance.