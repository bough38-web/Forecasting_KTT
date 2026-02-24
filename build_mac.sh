#!/bin/bash
echo "Building SalesExplorer for macOS..."

# Install dependencies if not present
pip install pyinstaller pandas openpyxl flask

# Clean previous builds
rm -rf build dist

# Run PyInstaller
pyinstaller sales_app.spec --clean

echo "Build complete. Check 'dist' directory for SalesExplorer.app"
