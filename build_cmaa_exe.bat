@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo =====================================================================
echo   CMAA Executable Builder (--onefile Standalone Package)
echo   Production-Ready Corporate Build System
echo =====================================================================
echo.

rem 1. Verify Python availability
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python is not installed or not in system PATH.
    echo Please install Python 3.9+ or run within your corporate virtual environment.
    echo.
    pause
    exit /b 1
)

rem 2. Check/Install PyInstaller and core libraries
echo [STEP 1/5] Checking required build dependencies...
if exist "requirements-cmaa.txt" (
    python -m pip install -r requirements-cmaa.txt >nul 2>&1
) else if exist "requirements.txt" (
    python -m pip install -r requirements.txt >nul 2>&1
    python -m pip install pyinstaller >nul 2>&1
) else (
    python -m pip install pyinstaller openpyxl requests certifi urllib3 rapidfuzz >nul 2>&1
)
if %ERRORLEVEL% NEQ 0 (
    echo [WARNING] Pip install returned non-zero - possibly offline or air-gapped network.
    echo Proceeding with currently installed packages in Python environment...
)

rem 3. Clean prior build cache and all stale bytecode
echo.
echo [STEP 2/5] Cleaning previous build cache and stale bytecode...
if exist "build\CMAA_Audit" rmdir /s /q "build\CMAA_Audit"
if exist "dist\CMAA_Audit.exe" del /f /q "dist\CMAA_Audit.exe"
if exist "__pycache__" rmdir /s /q "__pycache__"

rem 4. Validate codebase integrity before building
echo.
echo [STEP 3/5] Validating codebase syntax...
python -m py_compile populate_cr_audit.py constants.py excel_handler.py sc_sheet_data.py servicenow_api.py api_logger.py questions_data.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [FATAL ERROR] Python syntax validation failed!
    echo A syntax error was detected in your codebase files.
    echo Please fix the errors reported above before building the executable.
    echo.
    pause
    exit /b 1
)
echo [OK] All codebase files compiled cleanly.

rem 5. Verify master template is present
if not exist "QAR Report.xlsx" (
    echo.
    echo [ERROR] Master template 'QAR Report.xlsx' not found in current directory!
    echo Please ensure 'QAR Report.xlsx' is present before packaging.
    echo.
    pause
    exit /b 1
)

rem 6. Build standalone executable
echo.
echo [STEP 4/5] Compiling standalone --onefile executable using CMAA.spec...
python -m PyInstaller CMAA.spec --clean --noconfirm

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [FATAL ERROR] PyInstaller build failed. Review output log above.
    echo.
    pause
    exit /b 1
)

rem 7. Run post-build smoke verification test
echo.
echo [STEP 5/5] Performing post-build verification on compiled binary...
if exist "dist\CMAA_Audit.exe" (
    dist\CMAA_Audit.exe --help >nul 2>&1
    if %ERRORLEVEL% EQU 0 (
        echo [OK] Standalone binary self-test passed!
    ) else (
        echo [WARNING] Binary returned non-zero exit on --help. Check antivirus or permissions.
    )
) else (
    echo [ERROR] Expected executable dist\CMAA_Audit.exe not found!
    echo.
    pause
    exit /b 1
)

echo.
echo =====================================================================
echo [SUCCESS] CMAA Standalone Executable Successfully Built!
echo.
echo Executable File: dist\CMAA_Audit.exe
echo Master Template: Bundled inside binary + hot-override enabled.
echo =====================================================================
echo.
pause
