text

@echo off
title Portfolio — Install Requirements
color 0D
echo ================================================
echo   INSTALL REQUIREMENTS
echo ================================================
echo.
echo This will install the required Python library (openpyxl).
echo.

cd /d "%~dp0"

python --version
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ Python is NOT installed!
    echo.
    echo Please:
    echo   1. Go to https://python.org/downloads
    echo   2. Download Python 3.x (latest)
    echo   3. During install: CHECK "Add Python to PATH"
    echo   4. Re-run this file
    echo.
    pause
    exit /b 1
)

echo.
echo Installing openpyxl...
python -m pip install openpyxl

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ Install failed. Try running as Administrator.
    pause
    exit /b 1
)

echo.
echo ================================================
echo   ✅ All requirements installed!
echo   You can now use UPDATE_WEBSITE.bat
echo ================================================
pause
