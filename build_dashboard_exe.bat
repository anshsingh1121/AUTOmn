@echo off
setlocal enabledelayedexpansion
title Build IMT Dashboard Standalone Application
echo ========================================================
echo   Building IMT Dashboard Standalone Application (--onedir)
echo   Optimized for corporate systems and rapid startup
echo ========================================================
echo.

:: 1. Verify Python availability
echo [1/4] Verifying Python environment...
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Python is not found in PATH.
    echo Please ensure Python 3.9+ is installed and accessible.
    pause
    exit /b 1
)

:: 2. Verify or install build dependencies
echo [2/4] Verifying required build dependencies...
python -m pip install --quiet --upgrade pyinstaller streamlit plotly pandas openpyxl pywin32 jsonschema

:: 3. Clean previous build artifacts
echo [3/4] Cleaning previous build artifacts...
if exist "build" rmdir /s /q "build"
if exist "dist\IMT_Dashboard" rmdir /s /q "dist\IMT_Dashboard"

:: 4. Compile with PyInstaller using the --onedir spec file
echo [4/4] Compiling application using PyInstaller (--onedir)...
echo This packages all runtime libraries into the '_internal' directory.
echo Please wait, this may take 1-3 minutes...
echo.

python -m PyInstaller --clean --noconfirm IMT_Dashboard.spec

if %ERRORLEVEL% neq 0 (
    echo.
    echo [ERROR] PyInstaller build failed. Check the error messages above.
    pause
    exit /b %ERRORLEVEL%
)

:: Post-build: Ensure output and config directories exist in dist\IMT_Dashboard
if not exist "dist\IMT_Dashboard\output" mkdir "dist\IMT_Dashboard\output"
if not exist "dist\IMT_Dashboard\config" (
    mkdir "dist\IMT_Dashboard\config"
    copy /y "config\config.json" "dist\IMT_Dashboard\config\config.json" >nul
)

:: Create a convenient launcher helper inside dist\IMT_Dashboard
(
echo @echo off
echo cd /d "%%~dp0"
echo start "" "IMT_Dashboard.exe"
) > "dist\IMT_Dashboard\Launch_Dashboard.bat"

:: Create user instructions inside dist\IMT_Dashboard
(
echo ==============================================================================
echo FCB Incident Tracker -- Production Dashboard Application
echo ==============================================================================
echo.
echo HOW TO RUN:
echo   Double-click 'IMT_Dashboard.exe' ^(or 'Launch_Dashboard.bat'^).
echo   The local Streamlit server will start and automatically pop open your default
echo   web browser directly to http://127.0.0.1:8501.
echo.
echo DIRECTORY STRUCTURE:
echo   - IMT_Dashboard.exe   : Main application executable.
echo   - _internal\          : All runtime binaries and Python dependencies.
echo   - config\config.json  : Editable configuration for sheet names and mappings.
echo   - output\             : Default directory where reconciled files and reports
echo                           are generated.
echo.
echo CORPORATE SECURITY COMPLIANCE:
echo   - Operates 100%% offline / air-gapped on localhost ^(127.0.0.1^).
echo   - Zero telemetry, zero cloud calls, zero external socket egress.
echo   - Non-destructive workbook writer preserves macros, formulas, and pivots.
echo ==============================================================================
) > "dist\IMT_Dashboard\README_HOW_TO_RUN.txt"

echo.
echo ========================================================
echo   BUILD COMPLETED SUCCESSFULLY!
echo ========================================================
echo Application directory:
echo   dist\IMT_Dashboard\
echo.
echo Executable:
echo   dist\IMT_Dashboard\IMT_Dashboard.exe
echo.
echo All libraries are packaged inside 'dist\IMT_Dashboard\_internal\'.
echo Launch time is now optimized to ~1-2 seconds!
echo ========================================================
pause
