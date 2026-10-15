# Halo MCP server - Windows setup
# Run from this folder:  powershell -ExecutionPolicy Bypass -File .\setup.ps1

$ErrorActionPreference = "Stop"
$Dir = $PSScriptRoot
Set-Location $Dir
Get-ChildItem $Dir -File | Unblock-File -ErrorAction SilentlyContinue

Write-Host "`n== Halo MCP setup ==" -ForegroundColor Cyan

# 1. Find Python 3.10+
$pyCmd = $null; $pyArgs = @()
if (Get-Command py -ErrorAction SilentlyContinue) { $pyCmd = "py"; $pyArgs = @("-3") }
elseif (Get-Command python -ErrorAction SilentlyContinue) { $pyCmd = "python" }
if (-not $pyCmd) {
    Write-Host "Python 3.10+ not found. Install it from https://www.python.org/downloads/ (tick 'Add to PATH'), then re-run." -ForegroundColor Red
    exit 1
}
$ver = & $pyCmd @pyArgs -c "import sys; print(f'{sys.version_info[0]}.{sys.version_info[1]}')"
if ([version]$ver -lt [version]"3.10") { Write-Host "Python $ver found; 3.10+ is required." -ForegroundColor Red; exit 1 }
Write-Host "Python $ver found."

# 2. Virtual environment + packages
if (-not (Test-Path "$Dir\.venv")) { & $pyCmd @pyArgs -m venv "$Dir\.venv" }
$VenvPy = "$Dir\.venv\Scripts\python.exe"
& $VenvPy -m pip install --quiet --upgrade pip
& $VenvPy -m pip install --quiet -r "$Dir\requirements.txt"
Write-Host "Packages installed."

# 3. Halo details
Write-Host "`nFrom Halo: Configuration > Integrations > HaloPSA API" -ForegroundColor Cyan
$BaseUrl  = (Read-Host "Halo URL (e.g. https://yourco.halopsa.com)").Trim().TrimEnd('/')
$Tenant   = (Read-Host "Tenant (Halo-hosted only - leave blank if self-hosted)").Trim()
$ClientId = (Read-Host "Client ID").Trim()
$Scope    = (Read-Host "Token scope [all]").Trim(); if (-not $Scope) { $Scope = "all" }

# 4. Secret -> Windows Credential Manager (never written to a file)
Write-Host "`nPaste the Client Secret below (input is hidden). It is stored in Windows Credential Manager." -ForegroundColor Cyan
& $VenvPy -m keyring set halo-mcp halo_client_secret

# 5. Test
$env:HALO_BASE_URL = $BaseUrl; $env:HALO_TENANT = $Tenant; $env:HALO_CLIENT_ID = $ClientId; $env:HALO_SCOPE = $Scope
Write-Host "`nTesting connection..." -ForegroundColor Cyan
& $VenvPy "$Dir\server.py" --test
$ok = ($LASTEXITCODE -eq 0)

# 6. Config snippet for Claude desktop
$envBlock = [ordered]@{ HALO_BASE_URL = $BaseUrl; HALO_CLIENT_ID = $ClientId; HALO_SCOPE = $Scope }
if ($Tenant) { $envBlock.HALO_TENANT = $Tenant }
$snippet = [ordered]@{ mcpServers = [ordered]@{ halo = [ordered]@{
    command = $VenvPy; args = @("$Dir\server.py"); env = $envBlock } } }
$json = $snippet | ConvertTo-Json -Depth 6
$json | Out-File -Encoding utf8 "$Dir\claude_config_snippet.json"

Write-Host "`n== Add this to Claude desktop ==" -ForegroundColor Cyan
Write-Host "Claude desktop > Settings > Developer > Edit Config, merge in the 'halo' entry below"
Write-Host "(also saved to claude_config_snippet.json), then fully quit and reopen Claude.`n"
Write-Host $json
if (-not $ok) { Write-Host "`nThe connection test failed - fix the error above before adding it to Claude." -ForegroundColor Yellow }
