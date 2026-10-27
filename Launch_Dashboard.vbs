' ====================================================================
' FCB Incident Tracker - Silent Dashboard Launcher
' Double-clicking this script launches the Streamlit Dashboard
' directly in your browser without showing any black CMD window.
' ====================================================================

Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

' Get the directory of this script
CurrentDir = FSO.GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = CurrentDir

' Launch Streamlit silently (0 = hide window, False = don't wait)
WshShell.Run "python -m streamlit run app.py --server.headless=false", 0, False
