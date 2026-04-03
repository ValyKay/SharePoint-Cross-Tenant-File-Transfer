# ============================================================
#  SharePoint File Transfer -- Upgrade Script
#  Downloads the latest sp-transfer.ps1 and sp-transfer.bat
#  Your configuration (sp-transfer-config.ps1) is never touched
# ============================================================

$repoBase = "https://raw.githubusercontent.com/ValyKay/sharepoint-site-sync/main"
$scriptDir = $PSScriptRoot

$files = @("sp-transfer.ps1", "sp-transfer.bat", "sp-upgrade.ps1")

Clear-Host
Write-Host ""
Write-Host "  SharePoint File Transfer -- Upgrade" -ForegroundColor Cyan
Write-Host "  ====================================" -ForegroundColor Cyan
Write-Host ""

$anyFailed = $false

foreach ($file in $files) {
    $dest   = Join-Path $scriptDir $file
    $backup = "$dest.bak"
    $url    = "$repoBase/$file"

    # Back up existing file
    if (Test-Path $dest) {
        Copy-Item $dest $backup -Force
        Write-Host "  Backed up: $file -> $file.bak" -ForegroundColor Gray
    }

    # Download latest
    try {
        Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing -ErrorAction Stop
        Write-Host "  Updated:   $file" -ForegroundColor Green
    } catch {
        Write-Host "  FAILED:    $file -- $($_.Exception.Message)" -ForegroundColor Red
        # Restore from backup
        if (Test-Path $backup) {
            Copy-Item $backup $dest -Force
            Write-Host "  Restored:  $file from backup" -ForegroundColor Yellow
        }
        $anyFailed = $true
    }
}

Write-Host ""
if ($anyFailed) {
    Write-Host "  Upgrade completed with errors. See above." -ForegroundColor Yellow
} else {
    Write-Host "  Upgrade complete!" -ForegroundColor Green
}
Write-Host "  Your config (sp-transfer-config.ps1) was not modified." -ForegroundColor Gray
Write-Host ""
Read-Host "Press Enter to close"
