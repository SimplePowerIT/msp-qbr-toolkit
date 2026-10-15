# STR generator - Windows setup
# Run from this folder:  powershell -ExecutionPolicy Bypass -File .\setup.ps1

$ErrorActionPreference = "Stop"
$Dir = $PSScriptRoot
Set-Location $Dir
Get-ChildItem $Dir -File | Unblock-File -ErrorAction SilentlyContinue
Write-Host "`n== QBR STR generator setup ==" -ForegroundColor Cyan

# 1. Python: reuse the Halo connector's environment (it already has httpx, keyring, mcp)
$Py  = Join-Path (Split-Path $Dir -Parent) ".venv\Scripts\python.exe"
$PyW = Join-Path (Split-Path $Dir -Parent) ".venv\Scripts\pythonw.exe"
if (-not (Test-Path $Py)) { Write-Host "Can't find $Py - set up the Halo connector first." -ForegroundColor Red; exit 1 }

# 2. Quick local test (no SharePoint)
$test = (Read-Host "Client to test with first (e.g. ACME) - leave blank to skip").Trim()
if ($test) {
    & $Py "$Dir\run.py" --client $test
    if ($LASTEXITCODE -ne 0) { Write-Host "Local test failed - fix the error above, then re-run setup." -ForegroundColor Red; exit 1 }
    Write-Host "Local test OK. Files are in $Dir\output" -ForegroundColor Green
}

# 3. Microsoft Graph app registration details
Write-Host "`nMicrosoft Entra app registration (see README.md, step 1)" -ForegroundColor Cyan
$Tenant   = (Read-Host "Directory (tenant) ID").Trim()
$ClientId = (Read-Host "Application (client) ID").Trim()
$HostName = (Read-Host "SharePoint host [yourco.sharepoint.com]").Trim(); if (-not $HostName) { $HostName = "yourco.sharepoint.com" }
$SitePath = (Read-Host "Site path for the lists and client folders (blank = root site, or e.g. sites/Operations)").Trim()
$Library  = (Read-Host "Document library holding Clients\<Client> [SHAREPOINT]").Trim(); if (-not $Library) { $Library = "SHAREPOINT" }

Write-Host "`nPaste the app's client secret (hidden). It is stored in Windows Credential Manager." -ForegroundColor Cyan
& $Py -m keyring set str-generator graph_client_secret

$cfg = [ordered]@{ graph = [ordered]@{ tenant_id = $Tenant; client_id = $ClientId; hostname = $HostName;
                                      site_path = $SitePath; library = $Library; client_subfolder = "Strategic Tech Review" } }
$cfg | ConvertTo-Json -Depth 5 | Out-File -Encoding utf8 "$Dir\config.json"
Write-Host "Saved config.json"

# 4. Create the SharePoint lists and fill the client dropdown
& $Py "$Dir\run.py" --setup
if ($LASTEXITCODE -ne 0) { Write-Host "SharePoint setup failed - check the app permissions (README step 1)." -ForegroundColor Red; exit 1 }

# 5. Background task: start at sign-in, check for requests every minute
$TaskName = "QBR STR Generator"
$action   = New-ScheduledTaskAction -Execute $PyW -Argument "`"$Dir\run.py`" --poll" -WorkingDirectory $Dir
$trigger  = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERDOMAIN\$env:USERNAME"
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Days 0) `
            -RestartCount 5 -RestartInterval (New-TimeSpan -Minutes 5) -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Force | Out-Null
Start-ScheduledTask -TaskName $TaskName
Write-Host "`nBackground task '$TaskName' installed and started." -ForegroundColor Green
Write-Host "It runs while you're signed in. Log: $Dir\str-generator.log"
Write-Host "`nLast step: add the button to a SharePoint page (README step 3), using the button link printed above."
