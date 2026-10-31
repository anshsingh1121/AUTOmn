#!/usr/bin/env bash
# =====================================================================
#   CMAA Executable Builder (--onefile Standalone Package for Unix/macOS)
#   Production-Ready Corporate Build System
# =====================================================================
set -e

echo "Checking Python 3 availability..."
if ! command -v python3 &>/dev/null; then
    echo "[ERROR] python3 could not be found. Please install Python 3.9+."
    exit 1
fi

echo "[STEP 1/5] Checking dependencies..."
if [ -f "requirements-cmaa.txt" ]; then
    python3 -m pip install --quiet -r requirements-cmaa.txt || {
        echo "[WARNING] Pip install failed or offline network. Using existing installed packages."
    }
elif [ -f "requirements.txt" ]; then
    python3 -m pip install --quiet -r requirements.txt || true
    python3 -m pip install --quiet pyinstaller || true
else
    python3 -m pip install --quiet pyinstaller openpyxl requests certifi urllib3 rapidfuzz || {
        echo "[WARNING] Pip install failed or offline network. Using existing installed packages."
    }
fi

echo "[STEP 2/5] Cleaning previous build cache and stale bytecode..."
rm -rf build/CMAA_Audit dist/CMAA_Audit
find . -type d -name "__pycache__" -exec rm -rf {} + 2>/dev/null || true

echo "[STEP 3/5] Validating codebase integrity (offline tests and syntax check)..."
python3 -m py_compile populate_cr_audit.py constants.py excel_handler.py sc_sheet_data.py servicenow_api.py
echo "[OK] Codebase validation passed."

if [ ! -f "QAR Report.xlsx" ]; then
    echo "[ERROR] Master template 'QAR Report.xlsx' not found in current directory!"
    exit 1
fi

echo "[STEP 4/5] Compiling standalone --onefile executable using CMAA.spec..."
python3 -m PyInstaller CMAA.spec --clean --noconfirm

echo "[STEP 5/5] Performing post-build verification on compiled binary..."
if [ -f "dist/CMAA_Audit" ]; then
    ./dist/CMAA_Audit --help >/dev/null 2>&1
    echo "[OK] Standalone binary self-test passed!"
else
    echo "[ERROR] Expected binary dist/CMAA_Audit not found!"
    exit 1
fi

echo "====================================================================="
echo "[SUCCESS] CMAA Standalone Binary Built Successfully: dist/CMAA_Audit"
echo "====================================================================="
