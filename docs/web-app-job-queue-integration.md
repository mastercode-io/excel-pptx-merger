# Web App Integration Guide: Async Job Queue API

## Overview

This guide shows how to integrate web applications with the Excel-PowerPoint Merger's async job queue system to overcome the 10-second timeout limit for heavy Excel processing operations.

### Why Use the Job Queue?

- **Avoid Timeouts**: Heavy Excel processing can take 30-60+ seconds, exceeding typical web request timeouts
- **Better UX**: Provide real-time progress updates to users instead of hanging requests
- **Reliability**: Jobs continue processing even if the client disconnects temporarily
- **Scalability**: Handle multiple concurrent requests without blocking

### Supported Operations

- **`/api/v1/extract`**: Extract data from Excel files (SharePoint or upload)
- **`/api/v1/merge`**: Merge Excel data into PowerPoint templates
- **`/api/v1/update`**: Update Excel files with new data

## Quick Start

### 1. Start an Async Job

```javascript
const jobResponse = await fetch(`${API_BASE}/api/v1/jobs/start`, {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
  },
  body: JSON.stringify({
    endpoint: '/api/v1/extract',
    payload: {
      sharepoint_excel_url: 'https://your-sharepoint-file-url',
      sheet_names: ['Sheet1', 'Sheet2'],
      config: { /* your extraction config */ }
    }
  })
});

const { jobId } = await jobResponse.json();
```

### 2. Poll for Completion

```javascript
let result;
do {
  await new Promise(resolve => setTimeout(resolve, 2000)); // Wait 2 seconds

  const statusResponse = await fetch(`${API_BASE}/api/v1/jobs/${jobId}/status`);
  const status = await statusResponse.json();

  if (status.status === 'completed') {
    const resultResponse = await fetch(`${API_BASE}/api/v1/jobs/${jobId}/result`);
    result = await resultResponse.json();
    break;
  } else if (status.status === 'failed') {
    throw new Error(`Job failed: ${status.message}`);
  }

  // Optional: Update progress UI
  updateProgressBar(status.progress || 0);

} while (true);

console.log('Extraction complete:', result.data);
```

## Complete API Reference

### Base URL
```
Production: https://your-cloud-function-url.cloudfunctions.net
Development: http://localhost:8080
```

### Authentication
Currently no authentication required. Consider adding API keys for production use.

---

## POST /api/v1/jobs/start

Start a new async job for any supported endpoint.

### Request Headers
```
Content-Type: application/json
```

### Request Body
```typescript
{
  endpoint: '/api/v1/extract' | '/api/v1/merge' | '/api/v1/update',
  payload: {
    // Endpoint-specific parameters (see examples below)
  }
}
```

### Response (Immediate)
```json
{
  "success": true,
  "jobId": "job_1690123456789_abc123",
  "status": "started",
  "endpoint": "/api/v1/extract",
  "estimatedTime": "30-60 seconds"
}
```

### Status Codes
- **200**: Job created successfully
- **400**: Invalid request (bad endpoint, missing payload)
- **429**: Too many concurrent jobs (rate limited)
- **500**: Server error

---

## GET /api/v1/jobs/{jobId}/status

Check the current status of a running job.

### Response Examples

**Job Running:**
```json
{
  "success": true,
  "jobId": "job_1690123456789_abc123",
  "status": "running",
  "progress": 45,
  "message": "Processing Excel file...",
  "created_at": "2024-07-24T15:25:00Z",
  "updated_at": "2024-07-24T15:25:30Z"
}
```

**Job Completed:**
```json
{
  "success": true,
  "jobId": "job_1690123456789_abc123",
  "status": "completed",
  "progress": 100,
  "message": "Job completed successfully",
  "created_at": "2024-07-24T15:25:00Z",
  "updated_at": "2024-07-24T15:26:15Z"
}
```

**Job Failed:**
```json
{
  "success": true,
  "jobId": "job_1690123456789_abc123",
  "status": "failed",
  "progress": 30,
  "message": "SharePoint file not found",
  "created_at": "2024-07-24T15:25:00Z",
  "updated_at": "2024-07-24T15:25:45Z"
}
```

### Status Values
- `pending`: Job created, not started yet
- `running`: Currently processing
- `completed`: Successfully finished
- `failed`: Error occurred
- `expired`: Job too old (>5 minutes)

---

## GET /api/v1/jobs/{jobId}/result

Retrieve completed job results. **This endpoint deletes the job after retrieval.**

### Response (Success)
```json
{
  "success": true,
  "jobId": "job_1690123456789_abc123",
  "status": "completed",
  "data": {
    // Same structure as direct endpoint call
    "success": true,
    "message": "Data extracted successfully",
    "data": {
      "sheets": {
        "Sheet1": {
          "key_value_pairs": { "name": "John", "age": 30 },
          "tables": [...],
          "images": [...]
        }
      }
    }
  },
  "retrieved_at": "2024-07-24T15:30:00Z"
}
```

### Response (Job Not Ready)
```json
{
  "success": false,
  "jobId": "job_1690123456789_abc123",
  "status": "running",
  "error": "Job still processing",
  "message": "Use /api/v1/jobs/{jobId}/status to check progress"
}
```

---

## Endpoint-Specific Payloads

### Extract Job Payload
```json
{
  "endpoint": "/api/v1/extract",
  "payload": {
    "sharepoint_excel_url": "https://your-sharepoint-url",
    "sheet_names": ["Sheet1", "Sheet2"],
    "config": {
      "version": "1.0",
      "sheet_configs": {
        "Sheet1": {
          "subtables": [
            {
              "name": "basic_info",
              "type": "key_value_pairs",
              "header_search": {
                "method": "fixed_range",
                "range": "A1:B10"
              },
              "data_extraction": {
                "orientation": "horizontal"
              }
            }
          ]
        }
      }
    },
    "auto_detect": true,
    "max_rows": 1000
  }
}
```

### Merge Job Payload
```json
{
  "endpoint": "/api/v1/merge",
  "payload": {
    "sharepoint_excel_url": "https://your-excel-file",
    "sharepoint_pptx_url": "https://your-powerpoint-template",
    "config": { /* extraction config */ }
  }
}
```

### Update Job Payload
```json
{
  "endpoint": "/api/v1/update",
  "payload": {
    "sharepoint_excel_url": "https://your-excel-file",
    "update_data": {
      "Sheet1": {
        "basic_info": {
          "name": "Updated Name",
          "status": "Active"
        }
      }
    },
    "config": { /* update config */ }
  }
}
```

## JavaScript Implementation Examples

### Complete Extract Workflow
```javascript
class ExcelExtractor {
  constructor(baseUrl) {
    this.baseUrl = baseUrl.replace(/\/$/, ''); // Remove trailing slash
  }

  async extractData(sharePointUrl, sheetNames, config, onProgress = null) {
    try {
      // 1. Start the job
      const jobResponse = await fetch(`${this.baseUrl}/api/v1/jobs/start`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          endpoint: '/api/v1/extract',
          payload: {
            sharepoint_excel_url: sharePointUrl,
            sheet_names: sheetNames,
            config: config
          }
        })
      });

      if (!jobResponse.ok) {
        throw new Error(`Failed to start job: ${jobResponse.statusText}`);
      }

      const { jobId } = await jobResponse.json();
      console.log(`Started extraction job: ${jobId}`);

      // 2. Poll for completion
      return await this.pollForCompletion(jobId, onProgress);

    } catch (error) {
      console.error('Extraction failed:', error);
      throw error;
    }
  }

  async pollForCompletion(jobId, onProgress = null, maxAttempts = 150) {
    let attempts = 0;

    while (attempts < maxAttempts) {
      try {
        // Check status
        const statusResponse = await fetch(`${this.baseUrl}/api/v1/jobs/${jobId}/status`);

        if (!statusResponse.ok) {
          throw new Error(`Status check failed: ${statusResponse.statusText}`);
        }

        const status = await statusResponse.json();

        // Update progress callback
        if (onProgress) {
          onProgress({
            status: status.status,
            progress: status.progress || 0,
            message: status.message || 'Processing...'
          });
        }

        // Check completion
        if (status.status === 'completed') {
          // Get results
          const resultResponse = await fetch(`${this.baseUrl}/api/v1/jobs/${jobId}/result`);

          if (!resultResponse.ok) {
            throw new Error(`Failed to get results: ${resultResponse.statusText}`);
          }

          const result = await resultResponse.json();
          return result.data; // Return the actual extraction data
        }

        if (status.status === 'failed') {
          throw new Error(`Job failed: ${status.message}`);
        }

        if (status.status === 'expired') {
          throw new Error('Job expired - took too long to complete');
        }

        // Wait before next poll (exponential backoff)
        const waitTime = Math.min(2000 + (attempts * 100), 5000);
        await new Promise(resolve => setTimeout(resolve, waitTime));
        attempts++;

      } catch (error) {
        console.error(`Poll attempt ${attempts + 1} failed:`, error);

        if (attempts >= maxAttempts - 1) {
          throw new Error(`Polling failed after ${maxAttempts} attempts: ${error.message}`);
        }

        // Wait before retry
        await new Promise(resolve => setTimeout(resolve, 1000));
        attempts++;
      }
    }

    throw new Error(`Job did not complete within ${maxAttempts} attempts`);
  }
}

// Usage
const extractor = new ExcelExtractor('https://your-api-url');

const extractionConfig = {
  version: "1.0",
  sheet_configs: {
    "Sheet1": {
      subtables: [{
        name: "employee_data",
        type: "table",
        header_search: {
          method: "text_search",
          text: "Employee Information"
        },
        data_extraction: {
          orientation: "vertical",
          max_rows: 100
        }
      }]
    }
  }
};

extractor.extractData(
  'https://your-sharepoint-file-url',
  ['Sheet1'],
  extractionConfig,
  (progress) => {
    console.log(`Progress: ${progress.progress}% - ${progress.message}`);
    // Update your UI progress bar here
  }
).then(data => {
  console.log('Extraction completed:', data);
  // Process your extracted data
}).catch(error => {
  console.error('Extraction failed:', error);
  // Handle error in UI
});
```

### React Hook Example
```javascript
import { useState, useCallback } from 'react';

function useAsyncExtraction(baseUrl) {
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState(0);
  const [error, setError] = useState(null);
  const [data, setData] = useState(null);

  const extractData = useCallback(async (sharePointUrl, sheetNames, config) => {
    setLoading(true);
    setError(null);
    setProgress(0);
    setData(null);

    try {
      const extractor = new ExcelExtractor(baseUrl);

      const result = await extractor.extractData(
        sharePointUrl,
        sheetNames,
        config,
        (progressInfo) => {
          setProgress(progressInfo.progress);
        }
      );

      setData(result);
      setProgress(100);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  }, [baseUrl]);

  return { extractData, loading, progress, error, data };
}

// Component usage
function ExtractionComponent() {
  const { extractData, loading, progress, error, data } = useAsyncExtraction(
    'https://your-api-url'
  );

  const handleExtract = async () => {
    await extractData(
      'https://your-sharepoint-url',
      ['Sheet1'],
      { /* your config */ }
    );
  };

  return (
    <div>
      <button onClick={handleExtract} disabled={loading}>
        {loading ? 'Extracting...' : 'Extract Data'}
      </button>

      {loading && (
        <div>
          <div>Progress: {progress}%</div>
          <progress value={progress} max={100} />
        </div>
      )}

      {error && <div>Error: {error}</div>}

      {data && <div>Extraction completed! {JSON.stringify(data, null, 2)}</div>}
    </div>
  );
}
```

## Error Handling

### Common Errors and Solutions

| Error | Cause | Solution |
|-------|-------|----------|
| `400: Invalid endpoint` | Unsupported endpoint in job request | Use `/api/v1/extract`, `/api/v1/merge`, or `/api/v1/update` |
| `400: Missing payload` | Empty or invalid payload | Ensure payload contains required parameters |
| `400: SharePoint file not found` | Invalid SharePoint URL or permissions | Verify URL and SharePoint access |
| `429: Too many jobs` | Rate limit exceeded | Wait and retry, or implement queue on client |
| `404: Job not found` | Invalid job ID or job expired | Jobs expire after 10 minutes |
| `500: Server error` | Internal server error | Check server logs, retry after delay |

### Robust Error Handling Example
```javascript
async function robustExtraction(sharePointUrl, sheetNames, config, maxRetries = 3) {
  let lastError;

  for (let retry = 0; retry < maxRetries; retry++) {
    try {
      const extractor = new ExcelExtractor('https://your-api-url');
      return await extractor.extractData(sharePointUrl, sheetNames, config);

    } catch (error) {
      lastError = error;

      // Don't retry for these errors
      if (error.message.includes('SharePoint file not found') ||
          error.message.includes('Invalid configuration')) {
        throw error;
      }

      // Wait before retry (exponential backoff)
      if (retry < maxRetries - 1) {
        const waitTime = Math.pow(2, retry) * 1000; // 1s, 2s, 4s
        console.log(`Retry ${retry + 1}/${maxRetries} in ${waitTime}ms...`);
        await new Promise(resolve => setTimeout(resolve, waitTime));
      }
    }
  }

  throw new Error(`Failed after ${maxRetries} retries: ${lastError.message}`);
}
```

## Best Practices

### Polling Configuration
```javascript
const POLLING_CONFIG = {
  initialDelay: 1000,     // Start polling after 1 second
  normalInterval: 2000,   // Poll every 2 seconds normally
  maxInterval: 5000,      // Never poll more than every 5 seconds
  maxAttempts: 150,       // 5 minutes max (150 * 2s average)
  backoffMultiplier: 1.1  // Gradually increase interval
};
```

### Production Considerations

1. **Rate Limiting**: Implement client-side queuing if you need to process many files
2. **Error Logging**: Log all job failures with context for debugging
3. **User Feedback**: Always show progress and estimated completion time
4. **Timeout Handling**: Set reasonable timeouts and provide cancel functionality
5. **Resource Cleanup**: Always call `/result` endpoint to clean up completed jobs

### Performance Tips

1. **Batch Operations**: Group multiple sheet extractions in a single job
2. **Config Optimization**: Use specific ranges instead of auto-detection when possible
3. **Parallel Jobs**: You can run multiple jobs concurrently (up to rate limits)
4. **Caching**: Cache extraction configs and SharePoint file metadata

### Security Considerations

1. **SharePoint URLs**: Ensure URLs are validated and from trusted sources
2. **Config Validation**: Validate extraction configs on client side
3. **API Keys**: Consider implementing API key authentication for production
4. **Rate Limiting**: Implement client-side rate limiting to avoid 429 errors

## Debugging

### Enable Debug Logging
```javascript
class ExcelExtractor {
  constructor(baseUrl, debug = false) {
    this.baseUrl = baseUrl;
    this.debug = debug;
  }

  log(...args) {
    if (this.debug) {
      console.log('[ExcelExtractor]', ...args);
    }
  }

  async extractData(sharePointUrl, sheetNames, config, onProgress = null) {
    this.log('Starting extraction for:', sharePointUrl);
    this.log('Config:', JSON.stringify(config, null, 2));

    // ... rest of implementation with debug logging
  }
}
```

### Monitor Job Queue
```javascript
// Check overall job queue status
async function getJobStats(baseUrl) {
  const response = await fetch(`${baseUrl}/api/v1/jobs/stats`);
  return await response.json();
}

// List all your jobs
async function listJobs(baseUrl, status = null) {
  const url = status
    ? `${baseUrl}/api/v1/jobs?status=${status}`
    : `${baseUrl}/api/v1/jobs`;

  const response = await fetch(url);
  return await response.json();
}
```

## Testing

### Unit Test Example
```javascript
// Using Jest
describe('ExcelExtractor', () => {
  let extractor;
  let mockFetch;

  beforeEach(() => {
    mockFetch = jest.fn();
    global.fetch = mockFetch;
    extractor = new ExcelExtractor('https://test-api');
  });

  test('should start job and poll for completion', async () => {
    // Mock job start response
    mockFetch.mockResolvedValueOnce({
      ok: true,
      json: () => Promise.resolve({ jobId: 'test-job' })
    });

    // Mock status polling (running -> completed)
    mockFetch
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ status: 'running', progress: 50 })
      })
      .mockResolvedValueOnce({
        ok: true,
        json: () => Promise.resolve({ status: 'completed', progress: 100 })
      });

    // Mock result retrieval
    mockFetch.mockResolvedValueOnce({
      ok: true,
      json: () => Promise.resolve({
        data: { sheets: { Sheet1: { data: 'test' } } }
      })
    });

    const result = await extractor.extractData(
      'https://test-sharepoint-url',
      ['Sheet1'],
      { test: 'config' }
    );

    expect(result.sheets.Sheet1.data).toBe('test');
    expect(mockFetch).toHaveBeenCalledTimes(4);
  });
});
```

## Support

### Common Issues

1. **Jobs taking too long**: Check your extraction config - reduce max_rows or use specific ranges
2. **SharePoint access denied**: Ensure the service has proper permissions to the SharePoint site
3. **Memory errors**: Large files may hit memory limits - consider splitting into smaller jobs
4. **Network timeouts**: Implement proper retry logic with exponential backoff

### Getting Help

1. Check the server logs for detailed error messages
2. Test your SharePoint URLs directly in a browser first
3. Validate your extraction config against the schema
4. Use the debug mode to see detailed request/response information

For additional support, please refer to the main API documentation or contact your system administrator.
