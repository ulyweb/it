@echo off
title Windows 11 Maintenance Dashboard
color 0B

echo.
echo  ============================================
echo   Windows 11 Maintenance Dashboard
echo  ============================================
echo.
echo   Starting backend server...
echo   Dashboard will open in your browser.
echo.
echo   TIP: Right-click this file and choose
echo        "Run as Administrator" for full access
echo  ============================================
echo.

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo   ERROR: Python not found.
    echo   Install Python from https://www.python.org/downloads/
    echo   Make sure to check "Add to PATH" during install.
    pause
    exit /b 1
)

REM Run the launcher
python "%~dp0launcher.py"

pause
