#!/bin/bash

# Grant GCS permissions to Cloud Function service account
# Run this script with an account that has storage.buckets.setIamPolicy permission

set -e

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

echo -e "${BLUE}🔐 Granting GCS permissions to Cloud Function service account...${NC}"
echo "=================================================================="

# Configuration
BUCKET_NAME="excel-pptx-merger-storage-uk"
SERVICE_ACCOUNT="840912031336-compute@developer.gserviceaccount.com"
ROLE="roles/storage.objectAdmin"

# Check current account
CURRENT_ACCOUNT=$(gcloud config get-value account 2>/dev/null)
echo -e "${YELLOW}Current account: ${CURRENT_ACCOUNT}${NC}"
echo ""

# Option 1: Try to grant permissions
echo -e "${BLUE}Option 1: Granting Storage Object Admin role...${NC}"
echo "Command to run (requires bucket admin permissions):"
echo ""
echo -e "${GREEN}gcloud storage buckets add-iam-policy-binding gs://${BUCKET_NAME} \\
    --member=\"serviceAccount:${SERVICE_ACCOUNT}\" \\
    --role=\"${ROLE}\"${NC}"
echo ""

# Option 2: Alternative using gsutil
echo -e "${BLUE}Option 2: Alternative using gsutil...${NC}"
echo "Command to run:"
echo ""
echo -e "${GREEN}gsutil iam ch serviceAccount:${SERVICE_ACCOUNT}:objectAdmin gs://${BUCKET_NAME}${NC}"
echo ""

# Option 3: Create dedicated service account (recommended)
echo -e "${BLUE}Option 3: Create dedicated service account (recommended)...${NC}"
echo "Commands to run:"
echo ""
echo -e "${GREEN}# Create service account${NC}"
echo -e "${GREEN}gcloud iam service-accounts create excel-pptx-merger-sa \\
    --display-name=\"Excel-PPTX Merger Service Account\"${NC}"
echo ""
echo -e "${GREEN}# Grant storage permissions${NC}"
echo -e "${GREEN}gcloud projects add-iam-policy-binding 840912031336 \\
    --member=\"serviceAccount:excel-pptx-merger-sa@840912031336.iam.gserviceaccount.com\" \\
    --role=\"roles/storage.objectAdmin\"${NC}"
echo ""
echo -e "${GREEN}# Deploy function with new service account${NC}"
echo -e "${GREEN}gcloud functions deploy excel-pptx-merger \\
    --service-account=\"excel-pptx-merger-sa@840912031336.iam.gserviceaccount.com\" \\
    --region=europe-west2 \\
    --gen2${NC}"

echo ""
echo -e "${YELLOW}⚠️  Note: You need appropriate permissions to run these commands.${NC}"
echo -e "${YELLOW}   If you don't have permissions, ask your GCP admin to run them.${NC}"
echo ""
echo -e "${BLUE}Current issue: The Cloud Function cannot write to GCS bucket '${BUCKET_NAME}'${NC}"
echo -e "${BLUE}Service account: ${SERVICE_ACCOUNT}${NC}"