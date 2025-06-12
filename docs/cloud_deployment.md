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
gcloud storage buckets create gs://YOUR_BUCKET_NAME --location=us-central1
```

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
  --runtime=python39 \
  --region=us-central1 \
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
  --substitutions=_REGION=us-central1,_BUCKET_NAME=YOUR_BUCKET_NAME,_BASE_PREFIX=temp,_API_KEY=YOUR_API_KEY
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

- `https://REGION-PROJECT_ID.cloudfunctions.net/excel-pptx-merger`

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
  https://REGION-PROJECT_ID.cloudfunctions.net/excel-pptx-merger \
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

- View Cloud Function logs in the Google Cloud Console
- Monitor storage usage in the GCS bucket
- Check Cloud Function execution metrics for performance issues

## Security Considerations

- Keep your API key secure and rotate it regularly
- Consider implementing more robust authentication for production
- Set appropriate IAM permissions on your GCS bucket
- Configure appropriate timeout and memory settings for your workload
