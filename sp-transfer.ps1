# ============================================================
#  SharePoint Cross-Tenant File Transfer -- Differential Sync
#  Only transfers files that are new or modified since last run
#  Produces a per-session transfer report next to this script
#  Run manually by double-clicking sp-transfer.bat
#
#  Usage:
#    .\sp-transfer.ps1              # normal transfer
#    .\sp-transfer.ps1 -DryRun     # preview only, no files transferred
# ============================================================

param(
    [switch]$DryRun
)

# -- CONFIGURATION -- edit these values before first use -----
$site1Url       = "https://sharepoint-site-1-replace-this"
$site2Url       = "https://sharepoint-site-2-replace-this"

$site1Library   = "Shared Documents"   	# Library display name on Site1 - get from URL - this is the English generic
$site2Library   = "Shared Documents"    # Library display name on Site2 - get from URL - this is the English generic
# ------------------------------------------------------------

$sessionStamp   = Get-Date -Format "yyyy-MM-dd_HHmmss"
$tempPath       = "$env:TEMP\SPTransfer_$sessionStamp"
$debugLog       = "$env:TEMP\SPTransfer_debug.txt"
$reportFile     = Join-Path $PSScriptRoot "TransferReport_$sessionStamp.txt"

function Write-Log {
    param([string]$Message, [string]$Color = "White")
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
    Write-Host $line -ForegroundColor $Color
    Add-Content -Path $debugLog -Value $line
}

function Write-Report {
    param([string]$Line)
    Add-Content -Path $reportFile -Value $Line
}

# -- Check PnP.PowerShell is installed -----------------------
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host ""
    Write-Host "  PnP.PowerShell module is not installed." -ForegroundColor Red
    Write-Host "  Please run this command in PowerShell, then try again:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "      Install-Module PnP.PowerShell -Scope CurrentUser" -ForegroundColor Cyan
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

# -- Start ---------------------------------------------------
Clear-Host
Write-Host ""
if ($DryRun) {
    Write-Host "  SharePoint File Transfer  [DRY RUN]" -ForegroundColor Yellow
    Write-Host "  =====================================" -ForegroundColor Yellow
} else {
    Write-Host "  SharePoint File Transfer" -ForegroundColor Cyan
    Write-Host "  ========================" -ForegroundColor Cyan
}
Write-Host ""
Write-Log "Session started$(if ($DryRun) { ' (DRY RUN)' })"

New-Item -ItemType Directory -Path $tempPath | Out-Null

# -- Helper: get all files from a library using list items ---
# Returns objects with .ServerRelativeUrl, .RelativePath, .TimeLastModified
function Get-AllLibraryFiles {
    param([string]$LibraryName, $Connection)

    $results = [System.Collections.Generic.List[object]]::new()

    # Get root folder server-relative URL to compute relative paths
    $list = Get-PnPList -Identity $LibraryName -Connection $Connection -ErrorAction Stop
    $rootUrl = $list.RootFolder.ServerRelativeUrl.TrimEnd('/')

    $pageSize = 500
    $items = Get-PnPListItem -List $LibraryName -PageSize $pageSize `
             -Connection $Connection -ErrorAction Stop

    foreach ($item in $items) {
        # FSObjType: 0 = file, 1 = folder
        if ($item["FSObjType"] -ne 0) { continue }

        $serverUrl   = $item["FileRef"]
        $modified    = $item["Modified"]
        $relativePath = $serverUrl.Substring($rootUrl.Length).TrimStart('/')

        $results.Add([PSCustomObject]@{
            ServerRelativeUrl = $serverUrl
            RelativePath      = $relativePath
            TimeLastModified  = [datetime]$modified
        })
    }

    return $results
}

# -- Step 1: Connect to Site1 --------------------------------
Write-Host ""
Write-Log "Connecting to Site1 (your company)..." "Cyan"
Write-Host "  -> A browser window will open. Log in with your work account." -ForegroundColor Yellow
Write-Host ""

try {
    $conn1 = Connect-PnPOnline -Url $site1Url -UseWebLogin -ReturnConnection -ErrorAction Stop
} catch {
    Write-Log "ERROR: Could not connect to Site1. $_" "Red"
    Read-Host "Press Enter to exit"
    exit 1
}

$operator = (Get-PnPProperty -ClientObject (Get-PnPWeb -Connection $conn1) `
             -Property CurrentUser -Connection $conn1).LoginName -replace ".*\\|.*\|", ""

Write-Log "Connected to Site1 as: $operator" "Green"
Write-Log "Reading file list from library: $site1Library" "Green"

try {
    $site1Files = Get-AllLibraryFiles -LibraryName $site1Library -Connection $conn1
} catch {
    Write-Log "ERROR: Could not read library on Site1. Check library name in config. $_" "Red"
    Read-Host "Press Enter to exit"
    exit 1
}

if ($site1Files.Count -eq 0) {
    Write-Log "No files found in source library. Nothing to transfer." "Yellow"
    Read-Host "Press Enter to exit"
    exit 0
}

Write-Log "Found $($site1Files.Count) file(s) on Site1." "Green"

# -- Step 2: Connect to Site2 --------------------------------
Write-Host ""
Write-Log "Connecting to Site2 (customer site)..." "Cyan"
Write-Host "  -> A second browser window will open. Log in with the SAME work account." -ForegroundColor Yellow
Write-Host "  -> You may see a guest indication -- this is normal." -ForegroundColor Yellow
Write-Host ""

try {
    $conn2 = Connect-PnPOnline -Url $site2Url -UseWebLogin -ReturnConnection -ErrorAction Stop
} catch {
    Write-Log "ERROR: Could not connect to Site2. $_" "Red"
    Read-Host "Press Enter to exit"
    exit 1
}

# -- Step 3: Build Site2 inventory ---------------------------
Write-Log "Reading existing files on Site2 for comparison..." "Cyan"

$site2Index = @{}

try {
    $site2Files = Get-AllLibraryFiles -LibraryName $site2Library -Connection $conn2
    foreach ($f in $site2Files) {
        $site2Index[$f.RelativePath.ToLower()] = $f.TimeLastModified
    }
    Write-Log "Site2 has $($site2Files.Count) existing file(s)." "Green"
} catch {
    Write-Log "Target library on Site2 appears empty or does not exist yet. All files will be transferred." "Yellow"
}

# -- Step 4: Determine what needs transferring ---------------
$toTransfer = [System.Collections.Generic.List[object]]::new()
$skipped    = 0

foreach ($f in $site1Files) {
    $relativeKey = $f.RelativePath.ToLower()
    $status      = "NEW"

    if ($site2Index.ContainsKey($relativeKey)) {
        if ($f.TimeLastModified -le $site2Index[$relativeKey]) {
            $skipped++
            continue
        }
        $status = "UPDATED"
    }

    $toTransfer.Add([PSCustomObject]@{
        File         = $f
        RelativePath = $f.RelativePath
        Status       = $status
    })
}

Write-Log "To transfer: $($toTransfer.Count)  |  Unchanged (skipped): $skipped" "Cyan"

if ($toTransfer.Count -eq 0) {
    Write-Report "SharePoint Transfer Report"
    Write-Report "=========================="
    Write-Report "Date/Time : $(Get-Date -Format 'yyyy-MM-dd  HH:mm:ss')"
    Write-Report "Operator  : $operator"
    Write-Report ""
    Write-Report "No files required transfer -- everything already up to date."
    Write-Report "Files checked : $($site1Files.Count)"
    Write-Report ""

    Write-Host ""
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host "  Everything is already up to date." -ForegroundColor Green
    Write-Host "  Report saved to: $reportFile" -ForegroundColor Gray
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host ""
    Remove-Item $tempPath -Recurse -Force -ErrorAction SilentlyContinue
    Read-Host "Press Enter to close"
    exit 0
}

# -- Step 5: Transfer (or preview in dry-run mode) ----------

if ($DryRun) {
    # -- Dry-run: preview what would be transferred -------------
    Write-Host ""
    Write-Log "DRY RUN -- listing files that would be transferred:" "Yellow"

    $counter = 0
    foreach ($entry in $toTransfer) {
        $counter++
        $progress = "$counter/$($toTransfer.Count)"
        Write-Log "  [$progress] [$($entry.Status)] $($entry.RelativePath)" "Yellow"
    }

    # -- Write dry-run report -----------------------------------
    Write-Report "SharePoint Transfer Report  [DRY RUN]"
    Write-Report "======================================"
    Write-Report "Date/Time : $(Get-Date -Format 'yyyy-MM-dd  HH:mm:ss')"
    Write-Report "Operator  : $operator"
    Write-Report ""
    Write-Report "** No files were transferred -- this was a dry run **"
    Write-Report ""
    Write-Report "Files that would be transferred ($($toTransfer.Count))"
    Write-Report ("-" * 40)
    foreach ($e in $toTransfer) {
        $tag = "[$($e.Status)]".PadRight(10)
        Write-Report "  $tag $($e.RelativePath)"
    }
    Write-Report ""
    Write-Report "Files already up to date (would be skipped) : $skipped"
    Write-Report ""

    Remove-Item $tempPath -Recurse -Force -ErrorAction SilentlyContinue

    Write-Log "Dry run complete. Would transfer: $($toTransfer.Count) | Already up to date: $skipped"

    Write-Host ""
    Write-Host "  ============================================" -ForegroundColor Yellow
    Write-Host "  DRY RUN complete. No files were transferred." -ForegroundColor Yellow
    Write-Host "  Would transfer : $($toTransfer.Count) file(s)" -ForegroundColor Yellow
    Write-Host "  Already current: $skipped file(s)" -ForegroundColor Gray
    Write-Host "  Report         : $reportFile" -ForegroundColor Gray
    Write-Host "  ============================================" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to close"
    exit 0
}

# -- Actual transfer -----------------------------------------
Write-Host ""
Write-Log "Starting transfer..." "Cyan"

$transferredEntries = [System.Collections.Generic.List[object]]::new()
$failedEntries      = [System.Collections.Generic.List[string]]::new()
$uploaded = 0
$failed   = 0
$counter  = 0

foreach ($entry in $toTransfer) {
    $counter++
    $f            = $entry.File
    $relativePath = $entry.RelativePath
    $relativeDir  = Split-Path $relativePath -Parent
    $fileName     = Split-Path $relativePath -Leaf
    $localDir     = if ($relativeDir) { Join-Path $tempPath $relativeDir } else { $tempPath }
    $targetFolder = if ($relativeDir) { "$site2Library/$relativeDir" } else { $site2Library }
    $progress     = "$counter/$($toTransfer.Count)"

    if (-not (Test-Path $localDir)) {
        New-Item -ItemType Directory -Path $localDir | Out-Null
    }

    # Download from Site1
    try {
        Get-PnPFile -Url $f.ServerRelativeUrl `
                    -Path $localDir -Filename $fileName `
                    -AsFile -Force -Connection $conn1 -ErrorAction Stop
    } catch {
        Write-Log "  [$progress] FAILED download: $relativePath -- $_" "Red"
        $failedEntries.Add($relativePath)
        $failed++
        continue
    }

    # Upload to Site2
    $localFilePath = Join-Path $localDir $fileName
    try {
        Add-PnPFile -Path $localFilePath `
                    -Folder $targetFolder `
                    -Connection $conn2 -ErrorAction Stop
        Write-Log "  [$progress] [$($entry.Status)] $relativePath" "White"
        $transferredEntries.Add($entry)
        $uploaded++
    } catch {
        Write-Log "  [$progress] FAILED upload: $relativePath -- $_" "Red"
        $failedEntries.Add($relativePath)
        $failed++
    }

    Remove-Item $localFilePath -Force -ErrorAction SilentlyContinue
}

# -- Cleanup -------------------------------------------------
Remove-Item $tempPath -Recurse -Force -ErrorAction SilentlyContinue

# -- Write transfer report -----------------------------------
Write-Report "SharePoint Transfer Report"
Write-Report "=========================="
Write-Report "Date/Time : $(Get-Date -Format 'yyyy-MM-dd  HH:mm:ss')"
Write-Report "Operator  : $operator"
Write-Report ""

if ($transferredEntries.Count -gt 0) {
    Write-Report "Files transferred ($($transferredEntries.Count))"
    Write-Report ("-" * 40)
    foreach ($e in $transferredEntries) {
        $tag = "[$($e.Status)]".PadRight(10)
        Write-Report "  $tag $($e.RelativePath)"
    }
} else {
    Write-Report "Files transferred (0)"
    Write-Report ("-" * 40)
    Write-Report "  (none)"
}

Write-Report ""
Write-Report "Files skipped -- already up to date : $skipped"
Write-Report ""

if ($failedEntries.Count -gt 0) {
    Write-Report "Failed ($($failedEntries.Count))"
    Write-Report ("-" * 40)
    foreach ($path in $failedEntries) {
        Write-Report "  $path"
    }
} else {
    Write-Report "Failed (0)"
    Write-Report ("-" * 40)
    Write-Report "  (none)"
}

Write-Report ""

# -- Console summary -----------------------------------------
Write-Log "Transfer complete. Transferred: $uploaded | Skipped: $skipped | Failed: $failed"

Write-Host ""
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host "  Transfer complete." -ForegroundColor Green
Write-Host "  Transferred : $uploaded file(s)" -ForegroundColor Green
Write-Host "  Skipped     : $skipped file(s) (already up to date)" -ForegroundColor Gray
if ($failed -gt 0) {
    Write-Host "  Failed      : $failed file(s) -- see report for details" -ForegroundColor Red
}
Write-Host "  Report      : $reportFile" -ForegroundColor Gray
Write-Host "  ============================================" -ForegroundColor Cyan
Write-Host ""
Read-Host "Press Enter to close"
