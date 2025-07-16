#!/bin/bash

# Test script for /merge endpoint with JSON payload (base64 encoded files)

# Configuration
API_URL="http://localhost:5000/api/v1/merge"
EXCEL_FILE="tests/fixtures/sample_excel.xlsx"
PPTX_FILE="tests/fixtures/sample_template.pptx"
OUTPUT_FILE="merged_output.pptx"

# Colors for output
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

echo -e "${YELLOW}Testing /merge endpoint with JSON payload (base64 mode)${NC}"
echo "=================================================="

# Check if files exist
if [ ! -f "$EXCEL_FILE" ]; then
    echo -e "${RED}Error: Excel file not found: $EXCEL_FILE${NC}"
    exit 1
fi

if [ ! -f "$PPTX_FILE" ]; then
    echo -e "${RED}Error: PowerPoint file not found: $PPTX_FILE${NC}"
    exit 1
fi

echo "Input files:"
echo "  Excel: $EXCEL_FILE"
echo "  PowerPoint: $PPTX_FILE"
echo ""

# Convert files to base64
echo "Converting files to base64..."
EXCEL_BASE64=$(base64 -i "$EXCEL_FILE" | tr -d '\n')
PPTX_BASE64=$(base64 -i "$PPTX_FILE" | tr -d '\n')

# Show file sizes
echo "File sizes:"
echo "  Excel (base64): ${#EXCEL_BASE64} characters"
echo "  PowerPoint (base64): ${#PPTX_BASE64} characters"
echo ""

# Create JSON payload
JSON_PAYLOAD=$(cat <<EOF
{
  "excel_file": "$EXCEL_BASE64",
  "pptx_file": "$PPTX_BASE64",
  "excel_filename": "test_excel.xlsx",
  "pptx_filename": "test_template.pptx",
  "config": {
    "global_settings": {
      "image_extraction": {
        "enabled": true
      }
    }
  }
}
EOF
)

# Save payload to temp file (for debugging if needed)
TEMP_PAYLOAD="/tmp/merge_payload.json"
echo "$JSON_PAYLOAD" > "$TEMP_PAYLOAD"
echo "Payload saved to: $TEMP_PAYLOAD (for debugging)"
echo ""

# Test 1: Standard JSON request
echo -e "${YELLOW}Test 1: JSON request with application/json Content-Type${NC}"
echo "Sending request..."

RESPONSE=$(curl -s -w "\n%{http_code}" -X POST "$API_URL" \
  -H "Content-Type: application/json" \
  -d "$JSON_PAYLOAD" \
  -o "$OUTPUT_FILE")

HTTP_CODE=$(echo "$RESPONSE" | tail -n1)

if [ "$HTTP_CODE" = "200" ]; then
    echo -e "${GREEN}✓ Success! HTTP $HTTP_CODE${NC}"
    echo -e "${GREEN}✓ Merged file saved as: $OUTPUT_FILE${NC}"
    ls -lh "$OUTPUT_FILE"
else
    echo -e "${RED}✗ Failed! HTTP $HTTP_CODE${NC}"
    echo "Response:"
    cat "$OUTPUT_FILE"
    rm -f "$OUTPUT_FILE"
fi

echo ""

# Test 2: CRM compatibility mode (text/plain)
echo -e "${YELLOW}Test 2: JSON request with text/plain Content-Type (CRM mode)${NC}"
echo "Sending request..."

OUTPUT_FILE2="merged_output_crm.pptx"
RESPONSE2=$(curl -s -w "\n%{http_code}" -X POST "$API_URL" \
  -H "Content-Type: text/plain" \
  -d "$JSON_PAYLOAD" \
  -o "$OUTPUT_FILE2")

HTTP_CODE2=$(echo "$RESPONSE2" | tail -n1)

if [ "$HTTP_CODE2" = "200" ]; then
    echo -e "${GREEN}✓ Success! HTTP $HTTP_CODE2${NC}"
    echo -e "${GREEN}✓ CRM compatibility mode works!${NC}"
    echo -e "${GREEN}✓ Merged file saved as: $OUTPUT_FILE2${NC}"
    ls -lh "$OUTPUT_FILE2"
else
    echo -e "${RED}✗ Failed! HTTP $HTTP_CODE2${NC}"
    echo "Response:"
    cat "$OUTPUT_FILE2"
    rm -f "$OUTPUT_FILE2"
fi

echo ""
echo "=================================================="
echo -e "${GREEN}Tests completed!${NC}"

# Cleanup temp file
rm -f "$TEMP_PAYLOAD"

# Alternative: Create a minimal test with small files
echo ""
echo -e "${YELLOW}Tip: For Postman testing, you can:${NC}"
echo "1. Use the JSON from $TEMP_PAYLOAD (before it's deleted)"
echo "2. Or use this minimal example with small test files:"
echo ""
echo "Create minimal test files:"
echo "  echo 'test' > tiny.txt"
echo "  base64 -i tiny.txt"
echo ""
echo "Then in Postman:"
echo "- Set request type to POST"
echo "- URL: $API_URL"
echo "- Headers: Content-Type: application/json"
echo "- Body: raw JSON with the base64 content"