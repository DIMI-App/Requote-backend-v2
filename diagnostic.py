#!/usr/bin/env python3
"""
Day 13 Diagnostic Tool - Requote AI
This script helps diagnose issues with Document AI processing
"""

import os
import sys

print("=" * 70)
print("🔍 REQUOTE AI - DAY 13 DIAGNOSTIC TOOL")
print("=" * 70)

# Check 1: Environment Variables
print("\n📋 Step 1: Checking Environment Variables...")
env_vars = {
    'GOOGLE_APPLICATION_CREDENTIALS_JSON': os.getenv('GOOGLE_APPLICATION_CREDENTIALS_JSON'),
    'GCP_PROJECT_ID': os.getenv('GCP_PROJECT_ID', 'requote-ai-backend'),
    'GCP_LOCATION': os.getenv('GCP_LOCATION', 'eu'),
    'GCP_PROCESSOR_ID': os.getenv('GCP_PROCESSOR_ID', 'f02a4802c23ab664'),
    'OPENAI_API_KEY': os.getenv('OPENAI_API_KEY')
}

for key, value in env_vars.items():
    if value:
        if 'KEY' in key or 'JSON' in key:
            print(f"   ✅ {key}: SET (length: {len(value)})")
        else:
            print(f"   ✅ {key}: {value}")
    else:
        print(f"   ❌ {key}: NOT SET")

# Check 2: Directory Structure
print("\n📁 Step 2: Checking Directory Structure...")
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
directories = {
    'uploads': os.path.join(BASE_DIR, 'uploads'),
    'outputs': os.path.join(BASE_DIR, 'outputs')
}

for name, path in directories.items():
    if os.path.exists(path):
        files = os.listdir(path)
        print(f"   ✅ {name}/: EXISTS ({len(files)} files)")
        for f in files[:5]:  # Show first 5 files
            print(f"      • {f}")
    else:
        print(f"   ❌ {name}/: DOES NOT EXIST")
        os.makedirs(path, exist_ok=True)
        print(f"      ➜ Created directory")

# Check 3: Required Files
print("\n📄 Step 3: Checking Required Files...")
required_files = [
    'api.py',
    'test_process.py',
    'extract_items.py',
    'generate_offer_doc.py',
    'requirements.txt'
]

for filename in required_files:
    filepath = os.path.join(BASE_DIR, filename)
    if os.path.exists(filepath):
        size = os.path.getsize(filepath)
        print(f"   ✅ {filename}: EXISTS ({size} bytes)")
    else:
        print(f"   ❌ {filename}: MISSING")

# Check 4: Python Packages
print("\n📦 Step 4: Checking Python Packages...")
packages = [
    'flask',
    'google.cloud.documentai_v1',
    'openai',
    'docx'
]

for package in packages:
    try:
        __import__(package if '.' not in package else package.split('.')[0])
        print(f"   ✅ {package}: INSTALLED")
    except ImportError:
        print(f"   ❌ {package}: NOT INSTALLED")

# Check 5: Upload Status
print("\n📤 Step 5: Checking Upload Status...")
offer1_path = os.path.join(BASE_DIR, 'uploads', 'offer1.pdf')
if os.path.exists(offer1_path):
    size = os.path.getsize(offer1_path)
    print(f"   ✅ offer1.pdf: EXISTS ({size} bytes)")
else:
    print(f"   ❌ offer1.pdf: NOT FOUND")
    print(f"   ➜ Please upload a PDF file to test with")

# Check 6: Output Files
print("\n💾 Step 6: Checking Output Files...")
output_files = {
    'extracted_text.txt': os.path.join(BASE_DIR, 'outputs', 'extracted_text.txt'),
    'items_offer1.json': os.path.join(BASE_DIR, 'outputs', 'items_offer1.json'),
    'final_offer1.docx': os.path.join(BASE_DIR, 'outputs', 'final_offer1.docx')
}

for name, path in output_files.items():
    if os.path.exists(path):
        size = os.path.getsize(path)
        print(f"   ✅ {name}: EXISTS ({size} bytes)")
    else:
        print(f"   ⚠️  {name}: NOT YET CREATED")

# Summary
print("\n" + "=" * 70)
print("📊 DIAGNOSTIC SUMMARY")
print("=" * 70)

issues = []

if not env_vars['GOOGLE_APPLICATION_CREDENTIALS_JSON']:
    issues.append("Missing Google Cloud credentials")

if not env_vars['OPENAI_API_KEY']:
    issues.append("Missing OpenAI API key")

if not os.path.exists(offer1_path):
    issues.append("No PDF file uploaded for testing")

if issues:
    print("\n⚠️  ISSUES FOUND:")
    for i, issue in enumerate(issues, 1):
        print(f"   {i}. {issue}")
    print("\n💡 RECOMMENDED ACTIONS:")
    print("   1. Set missing environment variables in Render dashboard")
    print("   2. Redeploy the application")
    print("   3. Test with a PDF upload")
else:
    print("\n✅ All checks passed! System appears to be configured correctly.")
    print("   If you're still having issues, check the logs for API errors.")

print("=" * 70)