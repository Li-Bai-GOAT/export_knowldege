@echo off
setlocal
cd /d "%~dp0"

rem Full workflow: export, create _bge when enabled, upload when enabled.
rem The two write operations are controlled by config.json.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0start_hospital.ps1" -PauseOnExit

exit /b %ERRORLEVEL%
