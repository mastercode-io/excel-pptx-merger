#!/usr/bin/env python3
"""Test script to verify Cloud Function fix for /merge endpoint JSON mode."""

import base64
import json
import requests
import os

# Test configuration
LOCAL_URL = "http://localhost:5000/api/v1/merge"
CLOUD_URL = "https://us-central1-excel-pptx-merger.cloudfunctions.net/excel-pptx-merger/api/v1/merge"
EXCEL_FILE_PATH = "tests/fixtures/sample_excel.xlsx"
PPTX_FILE_PATH = "tests/fixtures/sample_template.pptx"

def create_test_payload():
    """Create a test payload that simulates CRM request."""
    
    # Check if test files exist
    if not os.path.exists(EXCEL_FILE_PATH):
        print(f"Error: Excel test file not found at {EXCEL_FILE_PATH}")
        return None
    
    if not os.path.exists(PPTX_FILE_PATH):
        print(f"Error: PowerPoint test file not found at {PPTX_FILE_PATH}")
        return None
    
    # Read and encode files as base64
    with open(EXCEL_FILE_PATH, "rb") as f:
        excel_b64 = base64.b64encode(f.read()).decode('utf-8')
    
    with open(PPTX_FILE_PATH, "rb") as f:
        pptx_b64 = base64.b64encode(f.read()).decode('utf-8')
    
    # Create payload that matches CRM format
    payload = {
        "excel_file": excel_b64,
        "pptx_file": pptx_b64,
        "excel_filename": "test_excel.xlsx",
        "pptx_filename": "test_template.pptx",
        "config": {
            "global_settings": {
                "image_extraction": {
                    "enabled": True
                }
            }
        }
    }
    
    return payload

def test_endpoint(url, payload, test_name):
    """Test an endpoint with the given payload."""
    print(f"\n=== {test_name} ===")
    print(f"URL: {url}")
    
    try:
        # Test with correct Content-Type
        response = requests.post(
            url,
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=120
        )
        
        print(f"Response status: {response.status_code}")
        
        if response.status_code == 200:
            print("✅ SUCCESS - JSON mode works!")
            print(f"Response content-type: {response.headers.get('content-type', 'Not specified')}")
            if 'application/vnd.openxmlformats-officedocument.presentationml.presentation' in response.headers.get('content-type', ''):
                print("✅ Response is a PowerPoint file")
            else:
                print("⚠️  Response is not a PowerPoint file")
        else:
            print(f"❌ FAILED - Status {response.status_code}")
            print(f"Response: {response.text}")
            
            # Try to parse error response
            try:
                error_data = response.json()
                if 'error' in error_data:
                    print(f"Error type: {error_data['error'].get('type', 'Unknown')}")
                    print(f"Error message: {error_data['error'].get('message', 'Unknown')}")
            except:
                pass
    
    except requests.exceptions.RequestException as e:
        print(f"❌ Request failed: {e}")

def main():
    print("Testing Cloud Function fix for /merge endpoint JSON mode")
    print("=" * 60)
    
    # Create test payload
    payload = create_test_payload()
    if not payload:
        print("Failed to create test payload - check file paths")
        return
    
    print(f"Payload created successfully:")
    print(f"  Excel file size (base64): {len(payload['excel_file'])} characters")
    print(f"  PowerPoint file size (base64): {len(payload['pptx_file'])} characters")
    
    # Test local endpoint first (if available)
    try:
        test_endpoint(LOCAL_URL, payload, "Local Test")
    except Exception as e:
        print(f"\n=== Local Test ===")
        print(f"❌ Local server not available: {e}")
    
    # Test Cloud Function endpoint
    test_endpoint(CLOUD_URL, payload, "Cloud Function Test")
    
    print("\n" + "=" * 60)
    print("Test completed!")
    print("\nIf Cloud Function test succeeds, the fix is working!")
    print("If it fails, check the Cloud Function logs for our enhanced debugging output.")

if __name__ == "__main__":
    main()