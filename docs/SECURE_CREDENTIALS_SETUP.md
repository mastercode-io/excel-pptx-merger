# Secure Graph API Credentials Setup

This document explains how to securely configure Microsoft Graph API credentials for production deployment using Google Cloud Secret Manager.

## Security Overview

- ✅ **config/graph_api.env** is in `.gitignore` - never committed to repository
- ✅ **Actual credentials** are stored in Google Secret Manager, not in code
- ✅ **cloudbuild.yaml** references secrets by name, not actual values
- ✅ **Environment variables** are injected securely during deployment

## Initial Setup (One-time)

### 1. Create Secrets in Google Secret Manager

Run these commands to store your Graph API credentials securely:

```bash
# Set your project ID
export PROJECT_ID="your-gcp-project-id"

# Create the secrets (replace with your actual values from Azure AD app registration)
echo "your-graph-client-id-uuid" | gcloud secrets create excel-pptx-merger-graph-client-id --data-file=-
echo "your-graph-client-secret-value" | gcloud secrets create excel-pptx-merger-graph-client-secret --data-file=-
echo "your-tenant-id-uuid" | gcloud secrets create excel-pptx-merger-graph-tenant-id --data-file=-
```

### 2. Grant Access to Cloud Build Service Account

```bash
# Get your project number
PROJECT_NUMBER=$(gcloud projects describe ${PROJECT_ID} --format="value(projectNumber)")

# Grant Secret Manager access to Cloud Build service account
gcloud secrets add-iam-policy-binding excel-pptx-merger-graph-client-id \
    --member="serviceAccount:${PROJECT_NUMBER}@cloudbuild.gserviceaccount.com" \
    --role="roles/secretmanager.secretAccessor"

gcloud secrets add-iam-policy-binding excel-pptx-merger-graph-client-secret \
    --member="serviceAccount:${PROJECT_NUMBER}@cloudbuild.gserviceaccount.com" \
    --role="roles/secretmanager.secretAccessor"

gcloud secrets add-iam-policy-binding excel-pptx-merger-graph-tenant-id \
    --member="serviceAccount:${PROJECT_NUMBER}@cloudbuild.gserviceaccount.com" \
    --role="roles/secretmanager.secretAccessor"
```

### 3. Grant Access to Cloud Functions Service Account

```bash
# Grant Secret Manager access to Cloud Functions default service account
gcloud secrets add-iam-policy-binding excel-pptx-merger-graph-client-id \
    --member="serviceAccount:${PROJECT_ID}@appspot.gserviceaccount.com" \
    --role="roles/secretmanager.secretAccessor"

gcloud secrets add-iam-policy-binding excel-pptx-merger-graph-client-secret \
    --member="serviceAccount:${PROJECT_ID}@appspot.gserviceaccount.com" \
    --role="roles/secretmanager.secretAccessor"

gcloud secrets add-iam-policy-binding excel-pptx-merger-graph-tenant-id \
    --member="serviceAccount:${PROJECT_ID}@appspot.gserviceaccount.com" \
    --role="roles/secretmanager.secretAccessor"
```

## Deployment Process

After the initial setup, normal deployments will automatically:

1. **Cloud Build** accesses secrets from Secret Manager during build
2. **Secrets are injected** as environment variables into the Cloud Function
3. **Application code** reads credentials from environment variables (same as before)
4. **No changes needed** to application code - it works exactly as before

## How It Works

### cloudbuild.yaml Configuration

```yaml
# Deploy step references secrets using $$ syntax
--set-env-vars=...,GRAPH_CLIENT_ID=$$GRAPH_CLIENT_ID,GRAPH_CLIENT_SECRET=$$GRAPH_CLIENT_SECRET,GRAPH_TENANT_ID=$$GRAPH_TENANT_ID
secretEnv: ['GRAPH_CLIENT_ID', 'GRAPH_CLIENT_SECRET', 'GRAPH_TENANT_ID']

# Secret Manager configuration
availableSecrets:
  secretManager:
  - versionName: projects/${PROJECT_ID}/secrets/excel-pptx-merger-graph-client-id/versions/latest
    env: 'GRAPH_CLIENT_ID'
  - versionName: projects/${PROJECT_ID}/secrets/excel-pptx-merger-graph-client-secret/versions/latest
    env: 'GRAPH_CLIENT_SECRET'
  - versionName: projects/${PROJECT_ID}/secrets/excel-pptx-merger-graph-tenant-id/versions/latest
    env: 'GRAPH_TENANT_ID'
```

### Application Code (No Changes Required)

The application code continues to work exactly as before:

```python
# This code remains unchanged
client_id = os.getenv("GRAPH_CLIENT_ID", "")
client_secret = os.getenv("GRAPH_CLIENT_SECRET", "")
tenant_id = os.getenv("GRAPH_TENANT_ID", "")
```

## Managing Secrets

### View Current Secrets

```bash
gcloud secrets list --filter="name:excel-pptx-merger-graph"
```

### Update a Secret

```bash
# Update client secret (example)
echo "new-client-secret-value" | gcloud secrets versions add excel-pptx-merger-graph-client-secret --data-file=-
```

### Delete a Secret (if needed)

```bash
# ⚠️ Use with caution - this will break deployments
gcloud secrets delete excel-pptx-merger-graph-client-secret
```

## Troubleshooting

### Build Fails with "Secret not found"

1. Verify secrets exist:
   ```bash
   gcloud secrets list --filter="name:excel-pptx-merger-graph"
   ```

2. Check IAM permissions:
   ```bash
   gcloud secrets get-iam-policy excel-pptx-merger-graph-client-id
   ```

### Application Can't Access Credentials

1. Check Cloud Function environment variables:
   ```bash
   gcloud functions describe excel-pptx-merger --region=europe-west2
   ```

2. Review application logs for detailed diagnostic information

## Security Benefits

1. **No Credentials in Code**: Actual values never appear in source code or version control
2. **Audit Trail**: Secret Manager logs all access to credentials
3. **Rotation**: Credentials can be updated without code changes
4. **Access Control**: Fine-grained IAM controls who can access secrets
5. **Encryption**: Secrets are encrypted at rest and in transit

## Fallback for Local Development

Local development continues to use `config/graph_api.env` file as before:

1. File is in `.gitignore` - never committed
2. Contains actual credentials for local testing
3. Environment variables take precedence if set

This ensures a smooth development experience while maintaining production security.