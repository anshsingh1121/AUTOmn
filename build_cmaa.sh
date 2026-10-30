#!/usr/bin/env bash
# =====================================================================
#   CMAA Executable Builder (--onefile Standalone Package for Unix/macOS)
# =====================================================================
set -e

echo "Checking Python 3 availability..."
if ! command -v python3 &>/dev/null; then
    echo "[ERROR] python3 could not be found. Please install Python 3.9+."
    exit 1
fi

echo "[STEP 1/3] Checking dependencies..."
python3 -m pip install --quiet pyinstaller openpyxl requests rapidfuzz || {
    echo "[WARNING] Pip install failed or offline network. Using existing installed packages."
}

echo "[STEP 2/3] Cleaning previous build cache..."
rm -rf build/CMAA_Audit dist/CMAA_Audit

echo "[STEP 3/3] Compiling standalone --onefile executable using CMAA.spec..."
python3 -m PyInstaller CMAA.spec --clean --noconfirm

echo "====================================================================="
echo "[SUCCESS] Build completed! Standalone binary: dist/CMAA_Audit"
echo "====================================================================="
