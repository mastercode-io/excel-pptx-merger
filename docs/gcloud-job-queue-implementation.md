# General Job Queue System Implementation

## Overview
This document outlines the implementation of a flexible, general-purpose job queue system that can make any existing endpoint asynchronous without modifying current implementations. This solves timeout issues by moving heavy processing to background threads while providing real-time progress tracking.

## Architecture
```
Client → Netlify Function → Zoho API (decode request_id)
                         ↓
Netlify Function → /jobs/start → Background handler processing
                         ↓
Netlify Function polls /jobs/{jobId}/status until complete
                         ↓
Netlify Function → /jobs/{jobId}/result → Return to client
```

## Key Features

### Universal Async Wrapper
- **Any Endpoint**: Can make `/api/v1/extract`, `/api/v1/merge`, or `/api/v1/update` asynchronous
- **No Code Changes**: Existing endpoints remain completely unchanged
- **Internal Processing**: Calls handler functions directly (not HTTP requests)
- **Flexible**: Easy to add support for new endpoints

### Job Management
- **Unique Job IDs**: Auto-generated with timestamp and UUID
- **Status Tracking**: Real-time progress and status updates
- **Result Storage**: Temporary storage with automatic cleanup
- **Rate Limiting**: Configurable concurrent job limits per client

## API Endpoints

### 1. POST /api/v1/jobs/start
**Purpose**: Start any endpoint as an async background job

**Request**:
```json
{
  "endpoint": "/api/v1/extract",
  "payload": {
    "sharepoint_excel_url": "https://...",
    "config": {
      "sheets": {
        "Sheet1": {
          "extraction_rules": [...]
        }
      }
    },
    "sheet_names": ["Sheet1"]
  }
}
```

**Response** (immediate):
```json
{
  "success": true,
  "jobId": "job_1690123456789_abc123",
  "status": "started",
  "endpoint": "/api/v1/extract",
  "estimatedTime": "30-60 seconds"
}
```

**Supported Endpoints**:
- `/api/v1/extract` - Extract data from Excel files
- `/api/v1/merge` - Merge Excel data into PowerPoint templates
- `/api/v1/update` - Update Excel files with new data

### 2. GET /api/v1/jobs/{jobId}/status
**Purpose**: Check the current status of a running job

**Response Examples**:
```json
// Job still processing
{
  "success": true,
  "jobId": "job_1690123456789_abc123",
  "status": "running",
  "progress": 45,
  "message": "Processing /api/v1/extract request...",
  "created_at": "2024-07-24T15:25:00Z",
  "updated_at": "2024-07-24T15:25:30Z"
}

// Job completed
{
  "success": true,
  "jobId": "job_1690123456789_abc123",
  "status": "completed",
  "progress": 100,
  "message": "Job completed successfully",
  "created_at": "2024-07-24T15:25:00Z",
  "updated_at": "2024-07-24T15:26:15Z"
}

// Job failed
{
  "success": true,
  "jobId": "job_1690123456789_abc123",
  "status": "failed",
  "progress": 30,
  "message": "Job failed: Excel file not found",
  "created_at": "2024-07-24T15:25:00Z",
  "updated_at": "2024-07-24T15:25:45Z"
}
```

**Status Values**:
- `pending` - Job created, not started
- `running` - Currently processing
- `completed` - Successfully finished
- `failed` - Error occurred
- `expired` - Job too old (>5 minutes)

### 3. GET /api/v1/jobs/{jobId}/result
**Purpose**: Retrieve completed job results and clean up storage

**Response** (success):
```json
{
  "success": true,
  "jobId": "job_1690123456789_abc123",
  "status": "completed",
  "data": {
    "success": true,
    "message": "Data extracted successfully",
    "data": {
      "sheets": {
        "Sheet1": {
          "key_value_pairs": {...},
          "tables": [...],
          "images": [...]
        }
      }
    }
  },
  "retrieved_at": "2024-07-24T15:30:00Z"
}
```

**Response** (job not ready):
```json
{
  "success": false,
  "jobId": "job_1690123456789_abc123",
  "status": "running",
  "error": "Job still processing",
  "message": "Use /api/v1/jobs/{jobId}/status to check progress"
}
```

**Important**: This endpoint **deletes the job from storage** after successful retrieval to prevent memory leaks.

### 4. GET /api/v1/jobs (Optional - Debug)
**Purpose**: List all jobs, optionally filtered by status

**Query Parameters**:
- `status` - Filter by job status (optional)

**Response**:
```json
{
  "success": true,
  "total_jobs": 5,
  "jobs": [
    {
      "id": "job_1690123456789_abc123",
      "endpoint": "/api/v1/extract",
      "status": "completed",
      "progress": 100,
      "created_at": "2024-07-24T15:25:00Z",
      "updated_at": "2024-07-24T15:26:15Z"
    }
  ]
}
```

### 5. DELETE /api/v1/jobs/{jobId} (Optional)
**Purpose**: Cancel/delete a job

**Response**:
```json
{
  "success": true,
  "message": "Job 'job_1690123456789_abc123' deleted successfully"
}
```

### 6. GET /api/v1/jobs/stats (Optional - Monitoring)
**Purpose**: Get job queue statistics

**Response**:
```json
{
  "success": true,
  "stats": {
    "total_jobs": 15,
    "by_status": {
      "running": 2,
      "completed": 10,
      "failed": 2,
      "pending": 1
    },
    "by_endpoint": {
      "/api/v1/extract": 8,
      "/api/v1/merge": 5,
      "/api/v1/update": 2
    }
  }
}
```

## Implementation Details

### Job Storage Structure
```python
# In-memory storage with thread safety
jobs = {
    "job_1690123456789_abc123": {
        "id": "job_1690123456789_abc123",
        "endpoint": "/api/v1/extract",
        "payload": {...},  # Full request payload
        "status": "running",
        "progress": 45,
        "created_at": "2024-07-24T15:25:00Z",
        "updated_at": "2024-07-24T15:25:30Z",
        "result": None,  # Populated when complete
        "error": None,
        "retry_count": 0
    }
}
```

### Background Processing Logic
```python
def process_job(job_id: str, handler_func: Callable):
    try:
        # Update status to running
        job.update_status(JobStatus.RUNNING, progress=10)

        # Call endpoint handler directly with payload
        result = handler_func(job.payload)  # e.g., extract_handler(payload)

        # Store result
        job_queue.complete_job(job_id, result)

    except Exception as e:
        # Handle errors
        job_queue.fail_job(job_id, str(e))
```

### Handler Function Mapping
```python
# Maps endpoints to their handler functions
handler_registry = {
    '/api/v1/extract': extract_handler,
    '/api/v1/merge': merge_handler,
    '/api/v1/update': update_handler
}
```

## Real-World Usage Example

### TMH Client App → Netlify → GCloud Workflow

1. **Client Request**: TMH app sends request_id to Netlify function
2. **Zoho API Call**: Netlify function calls Zoho API to decode request_id into full `/extract` request
3. **Start Job**: Netlify function calls `/api/v1/jobs/start`:
   ```javascript
   const response = await fetch(`${GCLOUD_FUNCTION_URL}/api/v1/jobs/start`, {
     method: 'POST',
     headers: { 'Content-Type': 'application/json' },
     body: JSON.stringify({
       endpoint: '/api/v1/extract',
       payload: decodedExtractRequest  // From Zoho API
     })
   });
   const { jobId } = await response.json();
   ```

4. **Poll Status**: Netlify function polls until complete:
   ```javascript
   let status;
   do {
     await new Promise(resolve => setTimeout(resolve, 2000)); // Wait 2 seconds
     const statusResponse = await fetch(`${GCLOUD_FUNCTION_URL}/api/v1/jobs/${jobId}/status`);
     status = await statusResponse.json();
   } while (status.status === 'running' || status.status === 'pending');
   ```

5. **Get Results**: Retrieve final data:
   ```javascript
   if (status.status === 'completed') {
     const resultResponse = await fetch(`${GCLOUD_FUNCTION_URL}/api/v1/jobs/${jobId}/result`);
     const result = await resultResponse.json();
     return result.data; // Return to client
   }
   ```

## Error Handling

### Retry Logic
- **Not Implemented**: Jobs run once - retries should be handled at the Netlify function level
- **Fail Fast**: Jobs fail immediately on errors for quick feedback

### Timeout Handling
- **Job Timeout**: 5 minutes maximum processing time
- **Auto-Cleanup**: Jobs older than 10 minutes are automatically deleted
- **Status Updates**: Clear error messages when jobs expire

### Rate Limiting
- **Concurrent Jobs**: Maximum 10 active jobs per client IP
- **Validation**: Endpoints must be explicitly allowed for async processing
- **Error Responses**: Clear messages when limits are exceeded

## Security Considerations

### Input Validation
- **Endpoint Whitelist**: Only allowed endpoints can be processed asynchronously
- **Payload Validation**: Job payloads are validated by the same logic as direct endpoint calls
- **Job ID Sanitization**: Job IDs use UUID format to prevent injection

### CORS Configuration
- **Netlify Domain**: Allow requests from Netlify function domain
- **Method Support**: POST for job creation, GET for status/results, DELETE for cancellation
- **Headers**: Proper CORS headers for all job queue endpoints

### Memory Management
- **Automatic Cleanup**: Completed jobs are deleted after result retrieval
- **Periodic Cleanup**: Background cleanup of expired jobs every hour
- **Memory Limits**: Consider Cloud Function memory limits for large result sets

## Deployment Integration

### Current Deployment Process
- **No Changes Required**: Job queue endpoints deploy automatically with existing process
- **Environment Variables**: Uses same configuration as existing endpoints
- **Secrets**: Leverages existing Secret Manager integration for SharePoint credentials

### CORS Configuration (Optional)
If needed, update Cloud Function to allow Netlify domain:
```python
# In main.py, add CORS headers for job endpoints
@app.after_request
def after_request(response):
    if request.path.startswith('/api/v1/jobs/'):
        response.headers.add('Access-Control-Allow-Origin', 'https://your-netlify-app.netlify.app')
        response.headers.add('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS')
        response.headers.add('Access-Control-Allow-Headers', 'Content-Type')
    return response
```

## Testing Examples

### Test Job Creation
```bash
curl -X POST https://your-gcloud-function.com/api/v1/jobs/start \
  -H "Content-Type: application/json" \
  -d '{
    "endpoint": "/api/v1/extract",
    "payload": {
      "sharepoint_excel_url": "https://...",
      "sheet_names": ["Sheet1"]
    }
  }'
```

### Test Status Check
```bash
curl https://your-gcloud-function.com/api/v1/jobs/job_1690123456789_abc123/status
```

### Test Result Retrieval
```bash
curl https://your-gcloud-function.com/api/v1/jobs/job_1690123456789_abc123/result
```

## Monitoring and Logging

### Job Lifecycle Logging
- **Job Creation**: Log when jobs are created with endpoint and client IP
- **Status Changes**: Log all status transitions (pending → running → completed/failed)
- **Result Retrieval**: Log when results are retrieved and cleaned up
- **Cleanup Events**: Log periodic cleanup operations

### Performance Monitoring
- **Job Duration**: Track how long different endpoint types take to process
- **Failure Rates**: Monitor success/failure rates by endpoint type
- **Memory Usage**: Track job storage memory consumption
- **Concurrent Jobs**: Monitor active job counts

## Benefits

### For Users
- **No Timeouts**: Heavy processing happens in background, no 10-second limits
- **Progress Tracking**: Real-time status updates and progress indicators
- **Reliability**: Failed jobs provide clear error messages
- **Flexibility**: Any endpoint can be made asynchronous

### For Developers
- **Zero Refactoring**: Existing endpoints work unchanged
- **Reusable**: Easy to add new endpoints to the job queue
- **Clean Architecture**: Clear separation between sync and async processing
- **Future-Proof**: Extensible design for additional features

This implementation provides a robust, scalable solution for handling long-running operations while maintaining the existing API contracts and adding powerful async capabilities.
