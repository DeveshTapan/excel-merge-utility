@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 (
    echo Python Launcher was not found.
    echo Install Python 3 from https://www.python.org/downloads/ and enable the Python Launcher.
    pause
    exit /b 1
)

py -3 append_merge_excel.py --gui
if errorlevel 1 (
    echo.
    echo The Excel Merge Utility closed with an error.
    pause
)

endlocal
