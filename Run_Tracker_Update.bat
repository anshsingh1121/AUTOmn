@echo off
title FCB Incident Tracker Automation
echo ===================================================
echo   FCB Incident Tracker Update Tool
echo ===================================================
echo.
echo Starting the automation pipeline...
echo.

:: Check if Python is installed
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Python is not installed or not in your PATH.
    echo Please install Python and try again.
    pause
    exit /b
)

:: Run the script
python -m src.main --production

echo.
echo ===================================================
echo   PROCESS FINISHED
echo ===================================================
echo.
pause
