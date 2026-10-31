@echo off
setlocal EnableDelayedExpansion
title RBI Notification Downloader - Setup
cd /d "%~dp0"

echo ============================================================
echo    RBI Notification Downloader - One-click Setup
echo ============================================================
echo.

REM ---------- 1. Python ----------
echo [1/6] Checking for Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo     Python not found. Installing Python 3.12 using winget...
    winget install -e --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements
    if errorlevel 1 (
        echo.
        echo     Automatic install did not work on this PC.
        echo     Opening the Python download page. Install it and TICK
        echo     "Add python.exe to PATH", then double-click this setup again.
        start https://www.python.org/downloads/
        pause
        exit /b 1
    )
    echo.
    echo     Python installed. CLOSE this window and double-click
    echo     setup_rbi_downloader.bat again to continue.
    pause
    exit /b 0
)
for /f "delims=" %%v in ('python --version') do echo     Found %%v
echo.

REM ---------- 2. Libraries ----------
echo [2/6] Installing required Python libraries...
python -m pip install --upgrade pip --quiet
python -m pip install requests beautifulsoup4 feedparser truststore pypdf --quiet
if errorlevel 1 (
    echo.
    echo     Library install failed. If your office uses a proxy, run:
    echo     python -m pip install --proxy http://PROXY:PORT requests beautifulsoup4 feedparser
    echo     (ask IT for the proxy address), then run this setup again.
    pause
    exit /b 1
)
echo     Libraries installed: requests, beautifulsoup4, feedparser, truststore, pypdf
echo.

REM ---------- 3. Script present? ----------
echo [3/6] Checking for rbi_downloader.py...
if not exist "%~dp0rbi_downloader.py" (
    echo     rbi_downloader.py was not found next to this setup file.
    echo     Put both files in the same folder ^(e.g. C:\Scripts^) and try again.
    pause
    exit /b 1
)
echo     Found %~dp0rbi_downloader.py
echo.

REM ---------- 4. Output folder ----------
echo [4/6] Checking S: drive...
if not exist "S:\" (
    echo     WARNING: S: drive is not available right now.
    echo     Connect/map the S: drive, then run this setup again.
    pause
    exit /b 1
)
if not exist "S:\RBI - Notifications" mkdir "S:\RBI - Notifications"
echo     Download folder ready: S:\RBI - Notifications
echo.

REM ---------- 5. Email settings ----------
echo [5/6] Opening the script in Notepad.
echo     Fill in SMTP_PASSWORD (and ANTHROPIC_API_KEY if you have one) in the CONFIG section,
echo     then SAVE and CLOSE Notepad to continue.
notepad "%~dp0rbi_downloader.py"
echo.

REM ---------- 6. Scheduled task ----------
echo [6/6] Creating scheduled task "RBI Alerts" (every 2 hours from 8 AM)...
for /f "delims=" %%p in ('python -c "import sys;print(sys.exec_prefix)"') do set "PYDIR=%%p"
set "PYW=!PYDIR!\pythonw.exe"
if not exist "!PYW!" set "PYW=!PYDIR!\python.exe"
schtasks /create /tn "RBI Alerts" /tr "\"!PYW!\" \"%~dp0rbi_downloader.py\"" /sc hourly /mo 2 /st 08:00 /f
if errorlevel 1 (
    echo     Could not create the task. Right-click this file and choose
    echo     "Run as administrator", then try again.
    pause
    exit /b 1
)
echo.

REM ---------- First run ----------
choice /c YN /m "Run the first download now (recommended; may take several minutes)"
if errorlevel 2 goto done
echo.
python "%~dp0rbi_downloader.py"

:done
echo.
echo ============================================================
echo  Setup complete.
echo  - Files:  S:\RBI - Notifications\FY ...\date\category\
echo  - Log:    S:\RBI - Notifications\rbi_downloader.log
echo  - Task:   "RBI Alerts" in Task Scheduler (every 2 hours)
echo ============================================================
pause
