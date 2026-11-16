@echo off
setlocal
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" goto install
where py >nul 2>&1
if errorlevel 1 (
    python -m venv .venv
) else (
    py -3 -m venv .venv
)
if errorlevel 1 goto failed
:install
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto failed
".venv\Scripts\python.exe" -m pip check
if errorlevel 1 goto failed
echo Dependencies ready. Use Start_Control_Center_Windows.bat to start.
echo TOI PDF export uses your installed desktop Microsoft Excel by default.
echo LibreOffice is optional. Install the workbook's Khmer fonts before exporting.
pause
exit /b 0
:failed
echo Installation failed. Read the error above. Python 3.10 or newer is required.
pause
exit /b 1
