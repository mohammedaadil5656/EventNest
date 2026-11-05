@echo off
title EventNest - Luxury Function Hall Booking System
color 0B
echo ===============================================================================
echo            EVENTNEST - FUNCTION HALL BOOKING & MANAGEMENT PLATFORM
echo ===============================================================================
echo.
echo  Executive Board:
echo    - Director & Co-Founder: Mohammed Aadil (+91 90352 80291)
echo    - CEO: Mr. Ahmad Talha Siddiqui
echo    - Operations Manager: Md Huzaif Masood
echo    - HR: Md Farhan Anwar
echo.
echo ===============================================================================
echo.

cd /d "%~dp0"

:: 1. Check Python installation
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python is not found or not in PATH!
    echo Please install Python 3.10+ from https://www.python.org/
    echo ** IMPORTANT: During installation, check "Add Python to PATH" **
    echo.
    pause
    exit /b 1
)

echo [1/4] Found Python:
python --version
echo.

:: 2. Check / Create Virtual Environment
if not exist "venv\Scripts\python.exe" (
    echo [2/4] Creating virtual environment (venv) for first-time setup...
    python -m venv venv
    if %ERRORLEVEL% NEQ 0 (
        echo [ERROR] Failed to create venv. Falling back to global Python.
    ) else (
        echo [OK] Virtual environment created successfully.
    )
) else (
    echo [2/4] Virtual environment (venv) already exists.
)
echo.

:: 3. Install Requirements
echo [3/4] Checking and installing required packages...
if exist "venv\Scripts\pip.exe" (
    ".\venv\Scripts\pip.exe" install -r requirements.txt --quiet
) else (
    pip install -r requirements.txt --quiet
)
echo [OK] All dependencies are ready.
echo.

:: 4. Launch Browser in Background
echo [4/4] Starting EventNest Server on http://127.0.0.1:5000 ...
start "" cmd /c "timeout /t 2 /nobreak >nul & start http://127.0.0.1:5000"

echo ===============================================================================
echo  WEBSITE IS RUNNING!
echo  Local URL: http://127.0.0.1:5000
echo.
echo  Default Logins:
echo   - Customer Demo: Mobile 9141928544 or Email mohammadmubasheer104@gmail.com (Pass: password)
echo   - Admin Panel:   Email mohammedaadil291@gamil.com (Pass: admin123)
echo.
echo  Press CTRL + C anytime to stop the server.
echo ===============================================================================
echo.

if exist "venv\Scripts\python.exe" (
    ".\venv\Scripts\python.exe" app.py
) else (
    python app.py
)

pause
