@echo off
echo Installing Required Libraries (PyInstaller, Pandas, OpenPyXL)...
python -m pip install pyinstaller pandas openpyxl

echo.
echo Building the executable...
python -m PyInstaller --onefile --console --name "FCB_Tracker_Updater" --add-data "config/config.json;config" --hidden-import openpyxl --hidden-import pandas src/main.py

echo.
echo =======================================================
echo Build complete!
echo The executable is located in the 'dist' folder.
echo You can run 'FCB_Tracker_Updater.exe' completely standalone,
echo or optionally place a 'config' folder next to it to override settings.
echo =======================================================
pause
