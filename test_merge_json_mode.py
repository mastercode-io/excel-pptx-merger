#!/usr/bin/env python3
"""Test script for JSON mode in /merge endpoint."""

import base64
import json
import requests
import os

# Test configuration
API_URL = "http://localhost:5000/api/v1/merge"
EXCEL_FILE_PATH = "tests/fixtures/sample_excel.xlsx"
PPTX_FILE_PATH = "tests/fixtures/sample_template.pptx"

def test_json_mode():
    """Test the /merge endpoint with JSON mode (base64 encoded files)."""
    
    # Check if test files exist
    if not os.path.exists(EXCEL_FILE_PATH):
        print(f"Error: Excel test file not found at {EXCEL_FILE_PATH}")
        return
    
    if not os.path.exists(PPTX_FILE_PATH):
        print(f"Error: PowerPoint test file not found at {PPTX_FILE_PATH}")
        return
    
    # Read and encode files as base64
    with open(EXCEL_FILE_PATH, "rb") as f:
        excel_b64 = base64.b64encode(f.read()).decode('utf-8')
    
    with open(PPTX_FILE_PATH, "rb") as f:
        pptx_b64 = base64.b64encode(f.read()).decode('utf-8')
    
    # Prepare JSON payload
    json_payload = {
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
    
    print("Testing JSON mode...")
    print(f"Excel file size (base64): {len(excel_b64)} characters")
    print(f"PowerPoint file size (base64): {len(pptx_b64)} characters")
    
    # Test 1: Standard JSON request
    print("\nTest 1: Standard JSON request with correct Content-Type")
    response = requests.post(
        API_URL,
        json=json_payload,
        headers={"Content-Type": "application/json"}
    )
    
    print(f"Response status: {response.status_code}")
    if response.status_code == 200:
        print("Success! File merged successfully")
        # Save the result
        with open("test_merged_json_mode.pptx", "wb") as f:
            f.write(response.content)
        print("Merged file saved as: test_merged_json_mode.pptx")
    else:
        print(f"Error: {response.text}")
    
    # Test 2: JSON request with text/plain Content-Type (CRM compatibility)
    print("\nTest 2: JSON request with text/plain Content-Type (CRM compatibility)")
    response = requests.post(
        API_URL,
        data=json.dumps(json_payload),
        headers={"Content-Type": "text/plain"}
    )
    
    print(f"Response status: {response.status_code}")
    if response.status_code == 200:
        print("Success! CRM compatibility mode works")
    else:
        print(f"Error: {response.text}")

def test_multipart_mode():
    """Test the /merge endpoint with traditional multipart mode."""
    
    print("\n\nTesting multipart mode...")
    
    # Check if test files exist
    if not os.path.exists(EXCEL_FILE_PATH):
        print(f"Error: Excel test file not found at {EXCEL_FILE_PATH}")
        return
    
    if not os.path.exists(PPTX_FILE_PATH):
        print(f"Error: PowerPoint test file not found at {PPTX_FILE_PATH}")
        return
    
    # Prepare multipart files
    with open(EXCEL_FILE_PATH, "rb") as excel_file:
        with open(PPTX_FILE_PATH, "rb") as pptx_file:
            files = {
                "excel_file": ("test_excel.xlsx", excel_file, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                "pptx_file": ("test_template.pptx", pptx_file, "application/vnd.openxmlformats-officedocument.presentationml.presentation")
            }
            
            data = {
                "config": json.dumps({
                    "global_settings": {
                        "image_extraction": {
                            "enabled": True
                        }
                    }
                })
            }
            
            response = requests.post(API_URL, files=files, data=data)
    
    print(f"Response status: {response.status_code}")
    if response.status_code == 200:
        print("Success! File merged successfully")
        # Save the result
        with open("test_merged_multipart_mode.pptx", "wb") as f:
            f.write(response.content)
        print("Merged file saved as: test_merged_multipart_mode.pptx")
    else:
        print(f"Error: {response.text}")

if __name__ == "__main__":
    print("Testing /merge endpoint dual mode support...")
    print("=" * 50)
    
    # Run both tests
    test_json_mode()
    test_multipart_mode()
    
    print("\n" + "=" * 50)
    print("Tests completed!")