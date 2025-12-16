@echo off
setlocal

REM === Settings ===
set "SCRIPT_PATH=C:\Temp\auto.ps1"
set "LOG_FILE=C:\Temp\run_log.txt"

REM === Check if Outlook is running using PowerShell ===
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
"if (-not (Get-Process OUTLOOK -ErrorAction SilentlyContinue)) { exit 1 }"

if errorlevel 1 (
    echo [%DATE% %TIME%] WARNING: Outlook is NOT running. Running fallback script. >> "%LOG_FILE%"
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_PATH%" -Fallback
    exit /b
)

REM === Check if Main Script Exists ===
if not exist "%SCRIPT_PATH%" (
    echo [%DATE% %TIME%] ERROR: Script not found: %SCRIPT_PATH% >> "%LOG_FILE%"
    exit /b 1
)

REM === Run main script ===
echo [%DATE% %TIME%] INFO: Running main script. >> "%LOG_FILE%"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_PATH%"
if errorlevel 1 (
    echo [%DATE% %TIME%] ERROR: Script failed. >> "%LOG_FILE%"
) else (
    echo [%DATE% %TIME%] SUCCESS: Script completed. >> "%LOG_FILE%"
)

echo Done.
