@echo off
setlocal enabledelayedexpansion

echo =====================================================================
echo   CMAA Executable Builder (--onefile Standalone Package)
echo =====================================================================
echo.

:: 1. Verify Python availability
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python is not installed or not in system PATH.
    echo Please install Python 3.9+ or run within your corporate virtual environment.
    pause
    exit /b 1
)

:: 2. Check/Install PyInstaller and core libraries
echo [STEP 1/3] Checking required build dependencies...
python -m pip install pyinstaller openpyxl requests rapidfuzz >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [WARNING] Pip install returned non-zero (possibly offline/air-gapped corporate network).
    echo Proceeding with locally installed packages...
)

:: 3. Clean prior build artifacts
echo.
echo [STEP 2/3] Cleaning previous build cache...
if exist "build\CMAA_Audit" rmdir /s /q "build\CMAA_Audit"
if exist "dist\CMAA_Audit.exe" del /f /q "dist\CMAA_Audit.exe"

:: 4. Build standalone executable
echo.
echo [STEP 3/3] Compiling standalone --onefile executable using CMAA.spec...
python -m PyInstaller CMAA.spec --clean --noconfirm

if %ERRORLEVEL% EQU 0 (
    echo.
    echo =====================================================================
    echo [SUCCESS] Build completed successfully!
    echo.
    echo Standalone Executable: dist\CMAA_Audit.exe
    echo.
    echo Features:
    echo   - Self-contained: All code, libraries, and QAR Report.xlsx bundled.
    echo   - Double-click: Runs interactive prompt with mode selection.
    echo   - CLI: Supports standard --change-nums and --start-date / --end-date flags.
    echo   - Output: Automatically creates report\ and api_logs\ in current directory.
    echo =====================================================================
) else (
    echo.
    echo [ERROR] Build failed. Review PyInstaller log output above.
)

echo.
pause
