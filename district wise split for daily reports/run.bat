@echo off
REM IFSC District Splitter - Windows launcher

echo.
echo ======================================================
echo    IFSC District Splitter
echo ======================================================
echo.

set "APP_PORT=5001"
set "APP_URL=http://localhost:%APP_PORT%"
set "FLASK_PORT=%APP_PORT%"
set "FLASK_DEBUG=false"

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Python is not installed or not in PATH
    echo Please install Python 3.8+ from https://www.python.org
    pause
    exit /b 1
)

echo Step 1: Creating folders...
if not exist "uploads" mkdir uploads
if not exist "downloads" mkdir downloads
echo Done.

echo.
echo ======================================================
echo Step 2: Starting IFSC District Splitter...
echo ======================================================
echo.
echo Opening %APP_URL% in your browser...
echo Press Ctrl+C to stop the server.
echo.

start "" /B powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command "Start-Sleep -Seconds 2; $null = [Diagnostics.Process]::Start('%APP_URL%')"
python app.py
pause
