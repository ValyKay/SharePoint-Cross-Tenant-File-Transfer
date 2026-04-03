@echo off
echo.
echo   SharePoint File Transfer
echo   ========================
echo.
echo   1) Normal sync
echo   2) Dry run (preview only)
echo   3) Upgrade script (keeps your config)
echo.
choice /C 123 /N /M "  Select mode (1, 2, or 3): "
if %errorlevel%==3 (
    PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0sp-upgrade.ps1"
) else if %errorlevel%==2 (
    PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0sp-transfer.ps1" -DryRun
) else (
    PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0sp-transfer.ps1"
)
