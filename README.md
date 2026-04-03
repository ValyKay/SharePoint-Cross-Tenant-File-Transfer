# SharePoint Cross-Tenant File Transfer

A PowerShell script for manually transferring files between two SharePoint document libraries across different tenants - including guest/external access scenarios.

## What it does

- Connects to a **source SharePoint site** (own tenant) and a **destination SharePoint site** (external/customer tenant where the user has guest contributor access)
- Performs a **differential sync**: only files that are new or modified since the last run are transferred - unchanged files are skipped
- Preserves **subfolder structure**
- Generates a **per-session transfer report** (plain text, suitable for client handover) saved next to the script files
- Maintains a **cumulative debug log** in `%TEMP%` for troubleshooting

## Requirements

- Windows with **Windows PowerShell 5.1** (built-in, no install needed)
- [PnP.PowerShell 1.12.0](https://www.powershellgallery.com/packages/PnP.PowerShell/1.12.0) - install once with:
  ```powershell
  Install-Module PnP.PowerShell -RequiredVersion 1.12.0 -Scope CurrentUser -Force
  ```
  > Version 1.12.0 specifically is required - newer versions (2.x, 3.x) require PowerShell 7 and will not be found by Windows PowerShell 5.1.
- Contributor access on the destination SharePoint site (guest/external invitation is sufficient)
- No admin rights needed for either installation or execution

## Files

| File | Purpose |
|------|---------|
| `sp-transfer.ps1` | Main script - configure this before first use |
| `sp-transfer.bat` | Launcher - users double-click this, choose normal sync or dry run |

Both files must be kept in the same folder.

## Configuration

Open `sp-transfer.ps1` and edit the block at the top:

```powershell
$site1Url      = "https://yourcompany.sharepoint.com/sites/Site1"
$site2Url      = "https://customertenant.sharepoint.com/sites/Site2"

$site1Library  = "Documents"   # Display name of the library on Site1
$site2Library  = "Documents"   # Display name of the library on Site2
```

> **Library names** must match the display name as shown in SharePoint, not the URL segment. Run `Get-PnPList | Select-Object Title, RootFolder` against the site to confirm the correct name.

## Usage

1. Double-click `sp-transfer.bat`
2. Select mode: **1** for normal sync, **2** for dry run (preview only)
3. Log in to Site1 in the browser window that opens
4. Log in to Site2 in the second browser window (same account, guest access)
5. Wait for the transfer to complete (or preview to finish)
6. Press Enter to close - the report is saved automatically

### Dry run

Dry run connects to both sites, compares files, and shows exactly what *would* be transferred - without downloading or uploading anything. Useful for verifying the diff before committing to a transfer.

```powershell
# Via the bat launcher: select option 2
sp-transfer.bat

# Or directly:
.\sp-transfer.ps1 -DryRun
```

The dry-run report is clearly marked `[DRY RUN]` and lists all files that would be transferred with their `[NEW]` or `[UPDATED]` status.

## Transfer report

Each run produces a timestamped report file in the same folder as the script:

```
TransferReport_2024-03-24_143210.txt
```

Example content:

```
SharePoint Transfer Report
==========================
Date/Time : 2024-03-24  14:32:10
Operator  : firstname.lastname

Files transferred (3)
----------------------------------------
  [NEW]      Data/survey_results.csv
  [UPDATED]  Reports/Q1_2024.xlsx
  [UPDATED]  Reports/Q2_2024.xlsx

Files skipped -- already up to date : 847

Failed (0)
----------------------------------------
  (none)
```

## Authentication

Uses `-UseWebLogin` (interactive browser popup) for both connections. This handles MFA and cross-tenant guest authentication without requiring app registrations or certificates. Tokens are not cached between runs - the user logs in fresh each time.

## Logs

| File | Content |
|------|---------|
| `TransferReport_<timestamp>.txt` | Per-session delivery record, saved next to the script |
| `%TEMP%\SPTransfer_debug.txt` | Cumulative debug log across all sessions, for IT troubleshooting |

## Notes

- Files are transferred one at a time: downloaded to a local temp folder, uploaded to Site2, then the local copy is deleted immediately. Peak local disk usage is one file at a time.
- The script is safe to run multiple times - `Add-PnPFile` overwrites the destination if the source is newer.
- The update warning (`A newer version of PnP PowerShell is available`) that appears on connect is harmless and can be ignored.
