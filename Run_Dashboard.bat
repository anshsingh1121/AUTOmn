@echo off
title IMT Incident Tracker Dashboard
echo ===================================================
echo   IMT Reconciliation & Executive Dashboard
echo ===================================================
echo.
echo Starting Streamlit Web Dashboard on port 8501...
echo.

:: Automatically open browser directly to port 8501
start "" http://127.0.0.1:8501

:: Run Streamlit with strict port 8501 binding
python -m streamlit run app.py --server.port 8501 --server.address 127.0.0.1 --browser.serverAddress 127.0.0.1 --browser.serverPort 8501

pause
