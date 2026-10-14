@echo off
title IMT Incident Tracker Dashboard
echo ===================================================
echo   IMT Reconciliation & Executive Dashboard
echo ===================================================
echo.
echo Starting Streamlit Web Dashboard...
echo A browser window will open automatically at http://localhost:8501
echo.

:: Check if Python is installed
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Python is not installed or not in your PATH.
    echo Please install Python and try again.
    pause
    exit /b
)

:: Run Streamlit Dashboard
python -m streamlit run app.py

echo.
echo ===================================================
echo   DASHBOARD CLOSED
echo ===================================================
echo.
pause
