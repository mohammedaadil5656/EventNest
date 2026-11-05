@echo off
title EventNest Luxury Function Hall Booking Server
color 0B
echo ===============================================================================
echo            EVENTNEST - FUNCTION HALL BOOKING & MANAGEMENT PLATFORM
echo ===============================================================================
echo.
echo  CEO: Mr. Ahmad Talha Siddiqui
echo  Director: Mohammed Aadil (Hotline: +91 90352 80291)
echo  Manager: Md Huzaif Masood
echo  HR: Md Farhan Anwar
echo.
echo  Starting EventNest local server on http://127.0.0.1:5000 ...
echo  Press CTRL+C to stop the server.
echo.
echo ===============================================================================

cd /d "%~dp0"
if exist ".\venv\Scripts\python.exe" (
    ".\venv\Scripts\python.exe" app.py
) else (
    echo Virtual environment not found. Running RUN_ON_ANOTHER_LAPTOP.bat ...
    call RUN_ON_ANOTHER_LAPTOP.bat
)

pause
