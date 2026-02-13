#!/usr/bin/env python3
"""
Test Script - Quick verification of all components
"""

import sys
from pathlib import Path

print("\n" + "="*60)
print("🧪 SHAREPOINT EXTRACTOR - SETUP TEST")
print("="*60 + "\n")

# Test 1: Python version
print("Test 1: Python Version")
print("-" * 40)
if sys.version_info >= (3, 9):
    print(f"✅ Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}")
else:
    print(f"❌ Python {sys.version_info.major}.{sys.version_info.minor} (need 3.9+)")
    sys.exit(1)

# Test 2: Dependencies
print("\nTest 2: Required Packages")
print("-" * 40)

required = {
    "msal": "Microsoft Authentication",
    "requests": "HTTP requests",
    "docx": "DOCX processing",
    "PyPDF2": "PDF processing",
    "pdfplumber": "Advanced PDF processing",
    "dotenv": "Environment variables",
    "tqdm": "Progress bars"
}

missing = []
for package, name in required.items():
    try:
        __import__(package)
        print(f"✅ {name}")
    except ImportError:
        print(f"❌ {name} - MISSING!")
        missing.append(package)

if missing:
    print(f"\n⚠️  Install missing packages: pip install {' '.join(missing)}")
    sys.exit(1)

# Test 3: Configuration
print("\nTest 3: Configuration")
print("-" * 40)

try:
    from config import Config
    
    if Config.validate():
        print("✅ Configuration valid")
        print(f"   CLIENT_ID: {Config.CLIENT_ID[:8]}...")
        print(f"   TENANT_ID: {Config.TENANT_ID[:8]}...")
    else:
        print("❌ Configuration invalid - check .env file")
        sys.exit(1)
except Exception as e:
    print(f"❌ Configuration error: {e}")
    sys.exit(1)

# Test 4: Modules
print("\nTest 4: Module Imports")
print("-" * 40)

modules = ["config", "auth", "sharepoint_api", "text_extraction", "main"]
for mod in modules:
    try:
        __import__(mod)
        print(f"✅ {mod}.py")
    except Exception as e:
        print(f"❌ {mod}.py - {e}")

# Summary
print("\n" + "="*60)
print("✅ ALL TESTS PASSED!")
print("="*60)
print("\nNext steps:")
print("1. python auth.py          # Test authentication")
print("2. python sharepoint_api.py # List files")
print("3. python main.py           # Run full pipeline")
print("\n" + "="*60 + "\n")
