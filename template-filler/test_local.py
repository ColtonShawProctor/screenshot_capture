#!/usr/bin/env python3
"""
Local test script for the template filler service.
Run this after starting the service locally.
"""

import requests
import json
import base64

# Test health endpoint
print("Testing health endpoint...")
response = requests.get("http://localhost:8000/health")
print(f"Health check: {response.json()}")

# Test fill endpoint with minimal data
print("\nTesting fill endpoint...")
fill_data = {
    "placeholders": {
        "LOAN_AMOUNT": "$25,000,000",
        "PROPERTY_ADDRESS": "89 Montauk Highway, East Moriches, NY",
        "LTC_RATIO": "102.1%",
        "DATE": "December 2024",
        "SPONSOR_NAME": "Test Sponsor LLC",
        "PROPERTY_LOCATION_DESCRIPTION": "East Moriches, Suffolk County, NY"
    },
    "output_filename": "test_output.docx"
}

response = requests.post(
    "http://localhost:8000/fill",
    json=fill_data
)

if response.status_code == 200:
    with open("test_output.docx", "wb") as f:
        f.write(response.content)
    print("✓ Fill endpoint test successful! Check test_output.docx")
else:
    print(f"✗ Fill endpoint test failed: {response.status_code}")
    print(response.text)

# Test fill-and-upload endpoint
print("\nTesting fill-and-upload endpoint...")
upload_data = {
    "placeholders": {
        "LOAN_AMOUNT": "$25,000,000",
        "PROPERTY_ADDRESS": "89 Montauk Highway, East Moriches, NY"
    },
    "output_key": "test/IDS_Test_Generated.docx"
}

response = requests.post(
    "http://localhost:8000/fill-and-upload",
    json=upload_data
)

if response.status_code == 200:
    print("✓ Fill-and-upload endpoint test successful!")
    print(json.dumps(response.json(), indent=2))
else:
    print(f"✗ Fill-and-upload endpoint test failed: {response.status_code}")
    print(response.text)