@echo off
title Build IMT Dashboard Executable
echo ========================================================
echo   Building Standalone IMT Dashboard Application (.exe)
echo ========================================================
echo.

:: 1. Ensure required build dependencies are installed
echo [1/2] Checking dependencies...
python -m pip install pyinstaller streamlit plotly pandas openpyxl

:: 2. Build the standalone executable
echo.
echo [2/2] Compiling IMT_Dashboard.exe with PyInstaller...
echo This will package Streamlit, Plotly, and the Dashboard UI into one executable.
echo Please wait, this process may take 3-5 minutes...
echo.

python -m PyInstaller --onefile --console --name "IMT_Dashboard" --collect-all streamlit --collect-all plotly --add-data "app.py;." --add-data "config/config.json;config" --add-data "src;src" run_dashboard.py

echo.
echo ========================================================
echo   BUILD FINISHED
echo ========================================================
echo The application executable is in the 'dist' folder:
echo dist\IMT_Dashboard.exe
echo.
echo Double-click 'dist\IMT_Dashboard.exe' to launch the Dashboard!
echo ========================================================
pause
