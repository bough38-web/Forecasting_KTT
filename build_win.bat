@echo off
echo Building SalesExplorer for Windows...

rem Install dependencies
pip install pyinstaller pandas openpyxl flask

rem Clean previous builds
if exist build rd /s /q build
if exist dist rd /s /q dist

rem Run PyInstaller
pyinstaller sales_app.spec --clean

echo Build complete. Check 'dist' directory for SalesExplorer.exe
pause
