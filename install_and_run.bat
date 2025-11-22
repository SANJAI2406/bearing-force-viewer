@echo off
echo ============================================
echo Bearing Force Viewer - Setup and Run
echo ============================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python from https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Installing required packages...
echo.

REM Install all required packages
pip install numpy matplotlib customtkinter pillow easyocr pandas openpyxl --quiet

if errorlevel 1 (
    echo.
    echo WARNING: Some packages may have failed to install.
    echo Trying alternative installation...
    pip install numpy matplotlib pillow pandas openpyxl --quiet
    pip install easyocr --quiet
)

echo.
echo ============================================
echo Starting Bearing Force Viewer...
echo ============================================
echo.

REM Run the application
python bearing_force_viewer.py

pause
