@echo off
setlocal EnableDelayedExpansion
title Windows 11 Maintenance Dashboard — Auto Setup
color 0B

:: ============================================================
::  Auto-elevate to Administrator if not already elevated
:: ============================================================
net session >nul 2>&1
if errorlevel 1 (
    echo  Requesting Administrator privileges...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

cls
echo.
echo  ============================================================
echo    Windows 11 Maintenance Dashboard  ^|  Auto-Setup
echo  ============================================================
echo.

:: ============================================================
::  STEP 1 — Check for Python
:: ============================================================
echo  [1/3] Checking for Python...

set PYTHON_OK=0
python --version >nul 2>&1
if not errorlevel 1 (
    set PYTHON_OK=1
    for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo        Found: %%v
)

if !PYTHON_OK!==0 (
    echo        Python not found. Auto-installing...
    call :install_python
    if errorlevel 1 (
        echo.
        echo  [FATAL] Python installation failed.
        echo          Please install manually: https://www.python.org/downloads/
        pause
        exit /b 1
    )
)

echo  [1/3] Python OK
echo.

:: ============================================================
::  STEP 2 — Verify launcher.py exists next to this .bat
:: ============================================================
echo  [2/3] Checking dashboard files...

if not exist "%~dp0launcher.py" (
    echo  [ERROR] launcher.py not found in: %~dp0
    echo          Make sure all 4 files are in the same folder.
    pause
    exit /b 1
)

if not exist "%~dp0dashboard.html" (
    echo  [ERROR] dashboard.html not found in: %~dp0
    pause
    exit /b 1
)

echo  [2/3] All files present
echo.

:: ============================================================
::  STEP 3 — Launch dashboard
:: ============================================================
echo  [3/3] Starting Maintenance Dashboard...
echo.
echo  ============================================================
echo    Server  : http://localhost:9191
echo    Browser : Opening automatically...
echo    Stop    : Press Ctrl+C in this window
echo  ============================================================
echo.

python "%~dp0launcher.py"

pause
exit /b 0


:: ============================================================
::  Subroutine: Download and silently install Python
:: ============================================================
:install_python
echo.
echo  -------------------------------------------------------
echo   Downloading Python 3.12 installer...
echo   (This may take a minute depending on your connection)
echo  -------------------------------------------------------
echo.

set PYTHON_URL=https://www.python.org/ftp/python/3.12.4/python-3.12.4-amd64.exe
set PYTHON_INSTALLER=%TEMP%\python_installer.exe

:: Use PowerShell to download (built into every Windows 11 system)
powershell -Command ^
  "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; " ^
  "$wc = New-Object System.Net.WebClient; " ^
  "$wc.DownloadFile('%PYTHON_URL%', '%PYTHON_INSTALLER%')"

if not exist "%PYTHON_INSTALLER%" (
    echo  [ERROR] Download failed. Check your internet connection.
    exit /b 1
)

echo  Download complete. Installing silently...
echo  (This takes 30-60 seconds, please wait...)
echo.

:: Silent install flags:
::   /quiet         — no GUI
::   InstallAllUsers=1 — install for all users
::   PrependPath=1  — add python to PATH automatically
::   Include_pip=1  — include pip
"%PYTHON_INSTALLER%" /quiet InstallAllUsers=1 PrependPath=1 Include_pip=1 Include_launcher=1

if errorlevel 1 (
    echo  [ERROR] Python installer exited with an error.
    del /f /q "%PYTHON_INSTALLER%" >nul 2>&1
    exit /b 1
)

del /f /q "%PYTHON_INSTALLER%" >nul 2>&1

echo  Python installed successfully!
echo.

:: Refresh PATH in this session so python is immediately usable
for /f "tokens=*" %%p in ('powershell -Command "[System.Environment]::GetEnvironmentVariable(\"PATH\",\"Machine\")"') do (
    set "PATH=%%p;%PATH%"
)

:: Verify install worked
python --version >nul 2>&1
if errorlevel 1 (
    echo  [WARNING] Python installed but PATH not refreshed yet.
    echo            Restarting script to pick up new PATH...
    :: Re-launch this same script — PATH will be correct on reopen
    start "" "%~f0"
    exit
)

exit /b 0
