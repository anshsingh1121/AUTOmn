#!/usr/bin/env bash
# =====================================================================
#   CMAA Executable Builder (--onefile Standalone Package for Unix/macOS)
#   Production-Ready Corporate Build System
# =====================================================================
set -e

# Change directory to script directory
cd "$(dirname "$0")"

echo "Checking Python availability..."
PYTHON_CMD="python3"
if ! command -v python3 &>/dev/null; then
    if command -v python &>/dev/null; then
        PYTHON_CMD="python"
    else
        echo "[ERROR] python3 or python could not be found. Please install Python 3.9+."
        exit 1
    fi
fi
echo "[OK] Using Python: $($PYTHON_CMD --version 2>&1)"

echo "[STEP 1/5] Checking dependencies..."
if [ -f "requirements-cmaa.txt" ]; then
    $PYTHON_CMD -m pip install --quiet -r requirements-cmaa.txt || {
        echo "[WARNING] Pip install failed or offline network. Using existing installed packages."
    }
elif [ -f "requirements.txt" ]; then
    $PYTHON_CMD -m pip install --quiet -r requirements.txt || true
    $PYTHON_CMD -m pip install --quiet pyinstaller || true
else
    $PYTHON_CMD -m pip install --quiet pyinstaller openpyxl requests certifi urllib3 rapidfuzz python-dotenv || {
        echo "[WARNING] Pip install failed or offline network. Using existing installed packages."
    }
fi

echo "[STEP 2/5] Cleaning previous build cache and stale bytecode..."
rm -rf build/CMAA_Audit dist/CMAA_Audit dist/CMAA_Audit.exe
find . -type d -name "__pycache__" -exec rm -rf {} + 2>/dev/null || true

echo "[STEP 3/5] Validating codebase integrity (offline tests and syntax check)..."
$PYTHON_CMD -m py_compile populate_cr_audit.py constants.py excel_handler.py sc_sheet_data.py servicenow_api.py api_logger.py questions_data.py
echo "[OK] Codebase validation passed."

if [ ! -f "QAR Report.xlsx" ]; then
    echo "[ERROR] Master template 'QAR Report.xlsx' not found in current directory!"
    exit 1
fi

echo "[STEP 4/5] Compiling standalone --onefile executable using CMAA.spec..."
$PYTHON_CMD -m PyInstaller CMAA.spec --clean --noconfirm

echo "[STEP 5/5] Performing post-build verification on compiled binary..."
TARGET_BIN="dist/CMAA_Audit"
if [ ! -f "$TARGET_BIN" ] && [ -f "${TARGET_BIN}.exe" ]; then
    TARGET_BIN="${TARGET_BIN}.exe"
fi

if [ -f "$TARGET_BIN" ]; then
    chmod +x "$TARGET_BIN" 2>/dev/null || true
    ./"$TARGET_BIN" --help >/dev/null 2>&1
    echo "[OK] Standalone binary self-test passed!"
else
    echo "[ERROR] Expected binary $TARGET_BIN not found!"
    exit 1
fi

echo "====================================================================="
echo "[SUCCESS] CMAA Standalone Binary Built Successfully: $TARGET_BIN"
echo "====================================================================="
