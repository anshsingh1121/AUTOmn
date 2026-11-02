# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller specification file for Change Management Audit Automation (CMAA).

Builds a standalone, corporate single-file executable (--onefile)
that bundles the Excel QAR template, all required dependencies,
interactive authentication prompt, enterprise input constraints
(max 100 Change Requests, max 1 month / 31-day date window),
and professional timestamped report nomenclature.
"""
import os
import sys

block_cipher = None

# Analysis collects all source files, bundled template assets, and dependencies
a = Analysis(
    ['populate_cr_audit.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('QAR Report.xlsx', '.'),
    ],
    hiddenimports=[
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.styles.fills',
        'openpyxl.styles.colors',
        'openpyxl.reader.excel',
        'openpyxl.writer.excel',
        'et_xmlfile',
        'requests',
        'urllib3',
        'certifi',
        'rapidfuzz',
        'thefuzz',
        'extract_msg',
        'constants',
        'excel_handler',
        'questions_data',
        'sc_sheet_data',
        'servicenow_api',
        'api_logger',
        'dotenv',
        'getpass',
        'shlex',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'scipy',
        'IPython',
        'jupyter',
        'pytest',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher,
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='CMAA_Audit',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
