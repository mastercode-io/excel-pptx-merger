#!/usr/bin/env python3
"""Generate a JSON payload for testing /merge endpoint in Postman."""

import base64
import json
import sys
import os

def generate_payload(excel_path, pptx_path, output_path="postman_payload.json"):
    """Generate a JSON payload with base64 encoded files."""
    
    # Check if files exist
    if not os.path.exists(excel_path):
        print(f"Error: Excel file not found: {excel_path}")
        sys.exit(1)
    
    if not os.path.exists(pptx_path):
        print(f"Error: PowerPoint file not found: {pptx_path}")
        sys.exit(1)
    
    # Read and encode files
    with open(excel_path, "rb") as f:
        excel_b64 = base64.b64encode(f.read()).decode('utf-8')
    
    with open(pptx_path, "rb") as f:
        pptx_b64 = base64.b64encode(f.read()).decode('utf-8')
    
    # Create payload
    payload = {
        "excel_file": excel_b64,
        "pptx_file": pptx_b64,
        "excel_filename": os.path.basename(excel_path),
        "pptx_filename": os.path.basename(pptx_path),
        "config": {
            "global_settings": {
                "image_extraction": {
                    "enabled": True
                }
            }
        }
    }
    
    # Save to file
    with open(output_path, "w") as f:
        json.dump(payload, f, indent=2)
    
    # Print stats
    print(f"✓ Payload generated successfully!")
    print(f"  Output file: {output_path}")
    print(f"  Excel size (base64): {len(excel_b64):,} characters")
    print(f"  PowerPoint size (base64): {len(pptx_b64):,} characters")
    print(f"  Total payload size: {os.path.getsize(output_path):,} bytes")
    print("")
    print("To use in Postman:")
    print("1. Set method to POST")
    print("2. URL: http://localhost:5000/api/v1/merge")
    print("3. Headers: Content-Type: application/json")
    print("4. Body: Select 'raw' and 'JSON'")
    print(f"5. Copy contents of {output_path} to body")
    print("")
    print("For smaller test payload, create minimal test files:")
    print("  echo 'test' > test.xlsx")
    print("  echo 'test' > test.pptx")
    print(f"  python {__file__} test.xlsx test.pptx minimal_payload.json")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python generate_postman_payload.py <excel_file> <pptx_file> [output_file]")
        print("")
        print("Example:")
        print("  python generate_postman_payload.py tests/fixtures/sample_excel.xlsx tests/fixtures/sample_template.pptx")
        sys.exit(1)
    
    excel_file = sys.argv[1]
    pptx_file = sys.argv[2]
    output_file = sys.argv[3] if len(sys.argv) > 3 else "postman_payload.json"
    
    generate_payload(excel_file, pptx_file, output_file)