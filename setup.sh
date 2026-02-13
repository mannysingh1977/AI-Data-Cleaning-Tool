#!/bin/bash

# Setup Script for SharePoint Document Extractor
# This script automates the initial setup process

echo ""
echo "============================================================"
echo "📚 SharePoint Document Extractor - Setup Script"
echo "============================================================"
echo ""

# Color codes
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Check Python version
echo "Step 1: Checking Python version..."
echo "----------------------------------------"
PYTHON_VERSION=$(python3 --version 2>&1 | awk '{print $2}')
PYTHON_MAJOR=$(echo $PYTHON_VERSION | cut -d. -f1)
PYTHON_MINOR=$(echo $PYTHON_VERSION | cut -d. -f2)

if [ "$PYTHON_MAJOR" -ge 3 ] && [ "$PYTHON_MINOR" -ge 10 ]; then
    echo -e "${GREEN}✅ Python $PYTHON_VERSION${NC}"
else
    echo -e "${RED}❌ Python $PYTHON_VERSION (need 3.10+)${NC}"
    echo "Please install Python 3.10 or higher"
    exit 1
fi

# Create virtual environment
echo ""
echo "Step 2: Creating virtual environment..."
echo "----------------------------------------"
if [ -d "venv" ]; then
    echo -e "${YELLOW}⚠️  Virtual environment already exists${NC}"
    read -p "Do you want to recreate it? (y/n) " -n 1 -r
    echo
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        rm -rf venv
        python3 -m venv venv
        echo -e "${GREEN}✅ Virtual environment recreated${NC}"
    else
        echo "Using existing virtual environment"
    fi
else
    python3 -m venv venv
    echo -e "${GREEN}✅ Virtual environment created${NC}"
fi

# Activate virtual environment
echo ""
echo "Step 3: Activating virtual environment..."
echo "----------------------------------------"
source venv/bin/activate
echo -e "${GREEN}✅ Virtual environment activated${NC}"

# Upgrade pip
echo ""
echo "Step 4: Upgrading pip..."
echo "----------------------------------------"
pip install --upgrade pip --quiet
echo -e "${GREEN}✅ pip upgraded${NC}"

# Install dependencies
echo ""
echo "Step 5: Installing dependencies..."
echo "----------------------------------------"
pip install -r requirements.txt --quiet
echo -e "${GREEN}✅ Dependencies installed${NC}"

# Create .env file if it doesn't exist
echo ""
echo "Step 6: Setting up environment variables..."
echo "----------------------------------------"
if [ -f ".env" ]; then
    echo -e "${YELLOW}⚠️  .env file already exists${NC}"
    echo "Skipping .env creation"
else
    cp .env.template .env
    echo -e "${GREEN}✅ .env file created from template${NC}"
    echo ""
    echo -e "${YELLOW}⚠️  IMPORTANT: You need to edit .env file and add:${NC}"
    echo "   - CLIENT_ID (from Azure Portal)"
    echo "   - TENANT_ID (from Azure Portal)"
    echo ""
    echo "Run this command to edit:"
    echo "  nano .env"
    echo "or"
    echo "  code .env"
fi

# Create directories
echo ""
echo "Step 7: Creating directories..."
echo "----------------------------------------"
mkdir -p downloaded_files
mkdir -p extracted_text
mkdir -p logs
echo -e "${GREEN}✅ Directories created${NC}"

# Run setup test
echo ""
echo "Step 8: Testing setup..."
echo "----------------------------------------"
python test_setup.py

echo ""
echo "============================================================"
echo "🎉 Setup Complete!"
echo "============================================================"
echo ""
echo "Next steps:"
echo "1. Edit .env file and add your Azure credentials:"
echo "   nano .env"
echo ""
echo "2. Test authentication:"
echo "   python auth.py"
echo ""
echo "3. Run the full pipeline:"
echo "   python main.py"
echo ""
echo "For detailed instructions, see:"
echo "  - QUICKSTART.md (quick reference)"
echo "  - README.md (full documentation)"
echo ""
echo "============================================================"
echo ""
