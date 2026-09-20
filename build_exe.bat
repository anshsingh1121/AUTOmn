@echo off
echo Installing Required Libraries (PyInstaller, Pandas, OpenPyXL)...
pip install pyinstaller pandas openpyxl

echo.
echo Building the executable...
pyinstaller --onefile --console --name "FCB_Tracker_Updater" --hidden-import openpyxl --hidden-import pandas src/main.py

echo.
echo =======================================================
echo Build complete!
echo The executable is located in the 'dist' folder.
echo You can move 'FCB_Tracker_Updater.exe' alongside your 'config' folder
echo to any Windows computer and run it without Python!
echo =======================================================
pause
