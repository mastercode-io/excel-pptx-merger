#!/bin/bash
echo "🚀 Deploying to Cloud Functions..."
gcloud functions deploy excel-pptx-merger \
  --gen2 \
  --runtime=python311 \
  --region=us-central1 \
  --source=. \
  --entry-point=excel_pptx_merger \
  --trigger-http \
  --allow-unauthenticated \
  --memory=1024MB \
  --timeout=540s \
  --set-env-vars="ENVIRONMENT=production,DEVELOPMENT_MODE=false,LOG_LEVEL=INFO,CLEANUP_TEMP_FILES=false,TEMP_FILE_RETENTION_SECONDS=3600,GOOGLE_CLOUD_BUCKET=excel-pptx-merger-storage" \
  --project=excel-pptx-merger

echo "✅ Deployment complete!"
echo "🔗 Testing health endpoint..."
curl "https://us-central1-excel-pptx-merger.cloudfunctions.net/excel-pptx-merger/api/v1/health"