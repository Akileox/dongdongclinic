@echo off
setlocal
title DongDong Clinic Report Generator - Portable

echo ==========================================
echo   DongDong Clinic Report Server Setup
echo ==========================================

:: 1. Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    py --version >nul 2>&1
    if %errorlevel% neq 0 (
        echo [ERROR] Python is not installed or not in PATH! 
        echo Please install Python from https://www.python.org/
        echo Make sure to check "Add Python to PATH" during installation.
        pause
        exit /b
    ) else (
        set PY_CMD=py
    )
) else (
    set PY_CMD=python
)

:: 2. Create virtual environment if it doesn't exist
if not exist ".venv" (
    echo [INFO] Creating virtual environment...
    %PY_CMD% -m venv .venv
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to create virtual environment. 
        pause
        exit /b
    )
)

:: 3. Activate virtual environment and install dependencies
echo [INFO] Activating virtual environment and checking dependencies...
if not exist ".venv\Scripts\activate.bat" (
    echo [ERROR] Virtual environment seems corrupted. Please delete .venv folder and try again.
    pause
    exit /b
)
call .venv\Scripts\activate.bat

:: Using a flag file to avoid re-installing every time
if not exist ".venv\installed_flag" (
    echo [INFO] Installing libraries (this may take a minute)...
    pip install --upgrade pip
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to install requirements.
        pause
        exit /b
    )
    echo [INFO] Installing Playwright browser...
    playwright install chromium
    if %errorlevel% neq 0 (
        echo [ERROR] Failed to install Playwright browsers.
        pause
        exit /b
    )
    echo done > .venv\installed_flag
)

:: 4. Start the server
echo ==========================================
echo   Server is starting! 
echo   Please go to: http://localhost:5000
echo ==========================================
python app.py
if %errorlevel% neq 0 (
    echo [ERROR] Server stopped unexpectedly.
    pause
)

