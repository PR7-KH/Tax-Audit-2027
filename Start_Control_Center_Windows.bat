@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Run Install_Update_Windows.bat first.
    pause
    exit /b 1
)
".venv\Scripts\python.exe" control_center.py
set "AUDIT_EXIT=%ERRORLEVEL%"
if not "%AUDIT_EXIT%"=="0" pause
exit /b %AUDIT_EXIT%
