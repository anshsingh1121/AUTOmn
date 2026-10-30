# -*- mode: python ; coding: utf-8 -*-
import sys
import os
from PyInstaller.utils.hooks import collect_all

datas = [
    ('app.py', '.'),
    ('config/config.json', 'config'),
    ('src', 'src'),
    ('.streamlit', '.streamlit'),
]
binaries = []
hiddenimports = [
    'win32com',
    'win32com.client',
    'pythoncom',
    'pywintypes',
    'openpyxl',
    'pandas',
    'numpy',
    'jinja2',
    'altair',
    'pyarrow',
    'src',
    'src.business_rules',
    'src.config_loader',
    'src.output_writer',
    'src.pivot_updater',
    'src.reconciler',
    'src.report_generator',
    'src.validator',
    'src.audit',
    'src.data_loader',
    'src.normalizer',
]

# Only streamlit and plotly have static HTML/JS frontend assets that require collect_all
for pkg in ['streamlit', 'plotly']:
    tmp_ret = collect_all(pkg)
    datas += tmp_ret[0]
    binaries += tmp_ret[1]
    hiddenimports += tmp_ret[2]

a = Analysis(
    ['run_dashboard.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'scipy', 'sqlite3', 'pandas.tests', 'numpy.tests'],
    noarchive=False,
    optimize=0,
)

# Exclude ucrtbase.dll from being copied into the dist folder because:
# 1. It is a core Windows System32 DLL present natively on every Windows 10/11 installation.
# 2. Antivirus and OneDrive cloud sync flag or lock ucrtbase.dll when written to user directories.
a.binaries = [x for x in a.binaries if not os.path.basename(x[0]).lower().startswith('ucrtbase')]

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='IMT_Dashboard',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='IMT_Dashboard',
)
