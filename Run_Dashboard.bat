@echo off
title IMT Incident Tracker Dashboard
cd /d "%~dp0"
echo ===================================================
echo   IMT Reconciliation & Executive Dashboard
echo ===================================================
echo.
echo Starting Dashboard on localhost...
echo Browser will automatically pop up when ready.
echo.

python run_dashboard.py

pause
