@echo off
setlocal enabledelayedexpansion

echo =====================================================================
echo   CMAA Executable Builder (--onefile Standalone Package)
echo   Production-Ready Corporate Build System
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
echo [STEP 1/5] Checking required build dependencies...
if exist "requirements-cmaa.txt" (
    python -m pip install -r requirements-cmaa.txt >nul 2>&1
) else (
    python -m pip install pyinstaller openpyxl requests certifi urllib3 rapidfuzz >nul 2>&1
)
if %ERRORLEVEL% NEQ 0 (
    echo [WARNING] Pip install returned non-zero (possibly offline/air-gapped network).
    echo Proceeding with currently installed packages in Python environment...
)

:: 3. Clean prior build cache and all stale __pycache__ bytecode
echo.
echo [STEP 2/5] Cleaning previous build cache and stale bytecode...
if exist "build\CMAA_Audit" rmdir /s /q "build\CMAA_Audit"
if exist "dist\CMAA_Audit.exe" del /f /q "dist\CMAA_Audit.exe"
for /d /r . %%d in (__pycache__) do @if exist "%%d" rd /s /q "%%d" 2>nul

:: 4. Validate codebase integrity before building
echo.
echo [STEP 3/5] Validating codebase integrity (offline tests and syntax check)...
python validate_reconstruction.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [FATAL ERROR] Code validation failed!
    echo A syntax error, broken import, or logic error was detected in your modified files.
    echo Please fix the errors reported above before building the executable.
    pause
    exit /b 1
)
echo [OK] All codebase files and audit logic validated successfully.

:: 5. Verify master template is present
if not exist "QAR Report.xlsx" (
    echo.
    echo [ERROR] Master template 'QAR Report.xlsx' not found in current directory!
    echo Please ensure 'QAR Report.xlsx' is present before packaging.
    pause
    exit /b 1
)

:: 6. Build standalone executable
echo.
echo [STEP 4/5] Compiling standalone --onefile executable using CMAA.spec...
python -m PyInstaller CMAA.spec --clean --noconfirm

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [FATAL ERROR] PyInstaller build failed. Review output log above.
    pause
    exit /b 1
)

:: 7. Run post-build smoke verification test
echo.
echo [STEP 5/5] Performing post-build verification on compiled binary...
if exist "dist\CMAA_Audit.exe" (
    dist\CMAA_Audit.exe --help >nul 2>&1
    if %ERRORLEVEL% EQU 0 (
        echo [OK] Standalone binary self-test passed!
    ) else (
        echo [WARNING] Standalone binary returned non-zero exit on --help. Check antivirus or permissions.
    )
) else (
    echo [ERROR] Expected executable dist\CMAA_Audit.exe not found!
    pause
    exit /b 1
)

echo.
echo =====================================================================
echo [SUCCESS] CMAA Standalone Executable Successfully Built!
echo.
echo Executable Location: dist\CMAA_Audit.exe
echo.
echo Production Features:
echo   - Standalone: All code, libraries, and QAR Report.xlsx bundled.
echo   - Hot Overrides: Can read external constants.py or .env if present.
echo   - Secure: API credentials masked (********) in logs.
echo   - Interactive & CLI: Double-click to run prompt; supports CLI flags.
echo   - Auto-Report: Generates report\ and api_logs\ in current directory.
echo =====================================================================
echo.
pause
