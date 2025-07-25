#!/bin/bash

# deploy.sh - Deploy excel-pptx-merger to Google Cloud Functions
# Direct deployment with manual secret management

set -e  # Exit on any error

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

echo -e "${BLUE}🚀 Deploying excel-pptx-merger to Google Cloud Functions...${NC}"
echo "=================================================================="

# Check if gcloud is installed and authenticated
if ! command -v gcloud &> /dev/null; then
    echo -e "${RED}❌ Error: gcloud CLI is not installed${NC}"
    exit 1
fi

# Get current project ID
PROJECT_ID=$(gcloud config get-value project 2>/dev/null)
if [ -z "$PROJECT_ID" ]; then
    echo -e "${RED}❌ Error: No GCP project set${NC}"
    echo "Please run: gcloud config set project YOUR_PROJECT_ID"
    exit 1
fi

echo -e "${BLUE}📋 Using GCP Project: ${PROJECT_ID}${NC}"

# Configuration
FUNCTION_NAME="excel-pptx-merger"
REGION="europe-west2"
RUNTIME="python312"
MEMORY="1024MB"
TIMEOUT="540s"
ENTRY_POINT="excel_pptx_merger"
GCS_BUCKET_NAME="excel-pptx-merger-storage"

# Secrets are managed manually
echo -e "${GREEN}✅ Using manually configured secrets${NC}"

# Deploy function directly
echo -e "\n${BLUE}🚀 Deploying function...${NC}"
gcloud functions deploy ${FUNCTION_NAME} \
    --gen2 \
    --runtime=${RUNTIME} \
    --region=${REGION} \
    --source=. \
    --entry-point=${ENTRY_POINT} \
    --trigger-http \
    --allow-unauthenticated \
    --memory=${MEMORY} \
    --timeout=${TIMEOUT} \
    --set-env-vars="STORAGE_BACKEND=LOCAL,SAVE_FILES=false,FLASK_DEBUG=false,DEVELOPMENT_MODE=false,LOG_LEVEL=INFO" \
    --set-secrets="GRAPH_CLIENT_ID=excel-pptx-merger-graph-client-id:latest,GRAPH_CLIENT_SECRET=excel-pptx-merger-graph-client-secret:latest,GRAPH_TENANT_ID=excel-pptx-merger-graph-tenant-id:latest"

# Get the function URL for testing
REGION="europe-west2"  # From cloudbuild.yaml
FUNCTION_URL="https://${REGION}-${PROJECT_ID}.cloudfunctions.net/excel-pptx-merger"

echo -e "\n${GREEN}✅ Deployment complete!${NC}"
echo "=================================================================="
echo -e "${GREEN}🔗 Function URL: ${FUNCTION_URL}${NC}"

# Test health endpoint
echo -e "\n${BLUE}🔍 Testing health endpoint...${NC}"
if curl -s "${FUNCTION_URL}/api/v1/health" | grep -q "healthy"; then
    echo -e "${GREEN}✅ Health check passed!${NC}"
else
    echo -e "${YELLOW}⚠️  Health check may have issues - check logs${NC}"
fi

echo -e "\n${BLUE}💡 Next steps:${NC}"
echo -e "${BLUE}   • Test SharePoint functionality: ${FUNCTION_URL}/api/v1/extract${NC}"
echo -e "${BLUE}   • Check logs: gcloud functions logs read excel-pptx-merger --region=${REGION}${NC}"
echo -e "${BLUE}   • Monitor performance in Cloud Console${NC}"

echo -e "\n${GREEN}🎉 Deployment completed successfully!${NC}"