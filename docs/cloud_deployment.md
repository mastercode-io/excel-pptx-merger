# Cloud Deployment Guide

This guide explains how to deploy the Excel-PowerPoint Merger as a Google Cloud Function with a public API endpoint.

## Prerequisites

Before deploying, ensure you have:

1. A Google Cloud Platform account
2. Google Cloud SDK installed locally
3. A Google Cloud Storage bucket for temporary file storage
4. Service account with appropriate permissions

## Setup Google Cloud Environment

### 1. Create a Google Cloud Storage Bucket

```bash
# Create a new bucket for temporary file storage
gcloud storage buckets create gs://YOUR_BUCKET_NAME --location=europe-west2
```

The application uses a standardized folder structure within the bucket that will be automatically created when the application runs.

## Temporary File Storage Structure

The application uses a standardized folder structure for temporary files in both local development and cloud deployment:

| Folder | Purpose |
|--------|---------|
| `/input` | Stores uploaded Excel and PowerPoint files |
| `/output` | Stores generated merged PowerPoint files |
| `/images` | Stores images extracted from Excel files |
| `/debug` | Stores debug information (in development mode) |

This structure is automatically created within the temporary directory or GCS bucket prefix.

### Google Cloud Storage Organization

When using GCS as the storage backend, files are organized as follows:

```
gs://YOUR_BUCKET_NAME/
└── temp/ (or your configured GCS_BASE_PREFIX)
    ├── input/
    │   ├── input_example.xlsx
    │   └── template_example.pptx
    ├── output/
    │   └── merged_example.pptx
    ├── images/
    │   └── extracted_image.png
    └── debug/
        └── example_debug_data.json
```

This organization helps maintain consistency between development and production environments and makes troubleshooting easier.

### 2. Create a Service Account

```bash
# Create a service account for the Cloud Function
gcloud iam service-accounts create excel-pptx-merger-sa \
  --display-name="Excel PowerPoint Merger Service Account"

# Grant the service account access to the bucket
gcloud storage buckets add-iam-policy-binding gs://YOUR_BUCKET_NAME \
  --member="serviceAccount:excel-pptx-merger-sa@YOUR_PROJECT_ID.iam.gserviceaccount.com" \
  --role="roles/storage.objectAdmin"
```

## Deployment Options

### Option 1: Manual Deployment

```bash
# Deploy the Cloud Function
gcloud functions deploy excel-pptx-merger \
  --gen2 \
  --runtime=python312 \
  --region=europe-west2 \
  --source=. \
  --entry-point=excel_pptx_merger \
  --trigger-http \
  --allow-unauthenticated \
  --memory=1024MB \
  --timeout=540s \
  --set-env-vars=STORAGE_BACKEND=GCS,GCS_BUCKET_NAME=YOUR_BUCKET_NAME,GCS_BASE_PREFIX=temp,API_KEY=YOUR_API_KEY
```

### Option 2: Using Cloud Build

```bash
# Deploy using Cloud Build
gcloud builds submit --config cloudbuild.yaml \
  --substitutions=_REGION=europe-west2,_BUCKET_NAME=YOUR_BUCKET_NAME,_BASE_PREFIX=temp,_API_KEY=YOUR_API_KEY
```

## Environment Variables

The following environment variables need to be configured:

| Variable | Description | Required |
|----------|-------------|----------|
| `STORAGE_BACKEND` | Storage backend to use (`LOCAL` or `GCS`) | Yes |
| `GCS_BUCKET_NAME` | Name of the Google Cloud Storage bucket | Yes (if using GCS) |
| `GCS_BASE_PREFIX` | Base prefix for objects in the bucket | No |
| `API_KEY` | API key for authentication | Yes |
| `GOOGLE_CLOUD_PROJECT` | Google Cloud project ID | Auto-detected |

## Using the Deployed API

### API Endpoints

Once deployed, your Cloud Function will expose the following endpoint:

- `https://europe-west2-PROJECT_ID.cloudfunctions.net/excel-pptx-merger`

### Authentication

Include your API key in the request:

```
Authorization: Bearer YOUR_API_KEY
```

Or as a query parameter:

```
?api_key=YOUR_API_KEY
```

### Example Request

```bash
curl -X POST \
  https://europe-west2-PROJECT_ID.cloudfunctions.net/excel-pptx-merger \
  -H "Authorization: Bearer YOUR_API_KEY" \
  -F "excel_file=@data.xlsx" \
  -F "pptx_file=@template.pptx" \
  -F "config={\"sheet_configs\":{...}}" \
  --output merged_output.pptx
```

## Local Testing with Cloud Storage

You can test the cloud storage functionality locally by:

1. Set up environment variables in a `.env` file based on `.env.example`
2. Set `STORAGE_BACKEND=GCS` and configure the GCS bucket
3. Set `GOOGLE_APPLICATION_CREDENTIALS` to point to your service account key file
4. Run the application locally

```bash
# Run locally with cloud storage
python -m src.main serve
```

## Monitoring and Troubleshooting

### Viewing Logs

To view Cloud Function logs in real-time:

```bash
gcloud functions logs read excel-pptx-merger --region=europe-west2
```

Or view logs in the Google Cloud Console:
1. Go to Cloud Functions in the Google Cloud Console
2. Click on your function name
3. Select the "Logs" tab

### Monitoring Performance

1. **Cloud Monitoring Dashboard**:
   - Go to Monitoring > Dashboards in Google Cloud Console
   - Create a custom dashboard for your function with metrics like:
     - Execution times
     - Memory usage
     - Invocation count
     - Error rate

2. **Setting Up Alerts**:
   - Create alerts for critical metrics:
     - Error rate exceeding threshold
     - Execution time above expected values
     - Memory usage approaching limit

### Storage Monitoring

Monitor your GCS bucket usage:

```bash
# List objects in your bucket
gcloud storage ls gs://YOUR_BUCKET_NAME/temp/

# Get bucket usage statistics
gcloud storage du -s gs://YOUR_BUCKET_NAME/
```

### Common Issues and Solutions

1. **Timeout Errors**: If processing large files causes timeouts, consider:
   - Increasing the timeout value (current: 540s)
   - Optimizing the code for faster processing

2. **Memory Errors**: If you encounter memory issues:
   - Increase memory allocation (current: 1024MB)
   - Implement more efficient memory management in the code

3. **Authentication Failures**: Check that:
   - The API key is correctly set in environment variables
   - The client is sending the API key properly in requests

## Security Considerations

- Keep your API key secure and rotate it regularly
- Consider implementing more robust authentication for production
- Set appropriate IAM permissions on your GCS bucket
- Configure appropriate timeout and memory settings for your workload
