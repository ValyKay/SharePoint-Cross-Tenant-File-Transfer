@echo off
echo.
echo   SharePoint File Transfer
echo   ========================
echo.
echo   1) Normal sync
echo   2) Dry run (preview only)
echo.
choice /C 12 /N /M "  Select mode (1 or 2): "
if %errorlevel%==2 (
    PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0sp-transfer.ps1" -DryRun
) else (
    PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0sp-transfer.ps1"
)