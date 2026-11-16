@echo off
setlocal
cd /d "%~dp0"
if not exist "%~dp0.venv-win\Scripts\python.exe" goto setup
"%~dp0.venv-win\Scripts\python.exe" -c "import sys,sysconfig; assert sys.version_info[:2]==(3,12) and sysconfig.get_platform()=='win-amd64'; import tkinter,openpyxl,xlsxwriter,pdfplumber,xlwings" >nul 2>&1
if errorlevel 1 goto setup
goto run
:setup
call "%~dp0Install_Update_Windows.bat" --no-pause
if errorlevel 1 (
    pause
    exit /b 1
)
:run
"%~dp0.venv-win\Scripts\python.exe" "%~dp0control_center.py"
set "AUDIT_EXIT=%ERRORLEVEL%"
if not "%AUDIT_EXIT%"=="0" pause
exit /b %AUDIT_EXIT%
