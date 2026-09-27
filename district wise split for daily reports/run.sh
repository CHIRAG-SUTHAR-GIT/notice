#!/bin/bash

# IFSC District Splitter - Bash Script

echo ""
echo "======================================================"
echo "   IFSC District Splitter - Setup and Run"
echo "======================================================"
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "Error: Python 3 is not installed"
    echo "Please install Python 3.8+ from https://www.python.org"
    exit 1
fi

echo "Step 1: Creating directories..."
mkdir -p uploads
mkdir -p downloads
mkdir -p templates
echo "Done."

echo ""
echo "Step 2: Installing requirements..."
python3 -m pip install -r requirements.txt
if [ $? -ne 0 ]; then
    echo "Error: Failed to install requirements"
    exit 1
fi
echo "Done."

echo ""
echo "======================================================"
echo "Step 3: Starting IFSC District Splitter..."
echo "======================================================"
echo ""
echo "Opening http://localhost:5000 in your browser..."
echo "Press Ctrl+C to stop the server."
echo ""

python3 app.py
