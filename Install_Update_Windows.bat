@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>&1
if errorlevel 1 (
    python "%~dp0setup_windows.py"
) else (
    py -3 "%~dp0setup_windows.py"
)
if errorlevel 1 goto failed
echo Setup complete. Use Start_Control_Center_Windows.bat to start.
if not "%~1"=="--no-pause" pause
exit /b 0
:failed
echo Read the setup error above. The supported Windows runtime is Python 3.12 x64.
if not "%~1"=="--no-pause" pause
exit /b 1
