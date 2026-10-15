# N-sight MCP server - Windows setup
# Run from this folder:  powershell -ExecutionPolicy Bypass -File .\setup.ps1

$ErrorActionPreference = "Stop"
$Dir = $PSScriptRoot
Set-Location $Dir
Get-ChildItem $Dir -File | Unblock-File -ErrorAction SilentlyContinue

Write-Host "`n== N-sight MCP setup ==" -ForegroundColor Cyan

# 1. Python environment: reuse the Halo connector's venv next door if it is there
$VenvPy = Join-Path (Split-Path $Dir -Parent) ".venv\Scripts\python.exe"
if (-not (Test-Path $VenvPy)) {
    $pyCmd = $null; $pyArgs = @()
    if (Get-Command py -ErrorAction SilentlyContinue) { $pyCmd = "py"; $pyArgs = @("-3") }
    elseif (Get-Command python -ErrorAction SilentlyContinue) { $pyCmd = "python" }
    if (-not $pyCmd) { Write-Host "Python 3.10+ not found. Install it from python.org (tick 'Add to PATH'), then re-run." -ForegroundColor Red; exit 1 }
    if (-not (Test-Path "$Dir\.venv")) { & $pyCmd @pyArgs -m venv "$Dir\.venv" }
    $VenvPy = "$Dir\.venv\Scripts\python.exe"
}
& $VenvPy -m pip install --quiet -r "$Dir\requirements.txt"
Write-Host "Using Python at $VenvPy"

# 2. Server for your region
Write-Host "`nN-sight API server for your region:" -ForegroundColor Cyan
Write-Host "  1) Americas        www.am.remote.management"
Write-Host "  2) United States   www.systemmonitor.us"
Write-Host "  3) United Kingdom  www.systemmonitor.co.uk"
Write-Host "  4) Europe          wwweurope1.systemmonitor.eu.com"
Write-Host "  5) Australia       www.system-monitor.com"
Write-Host "  6) Asia            wwwasia.system-monitor.com"
Write-Host "  Or type the server name shown with your API key."
$choice = (Read-Host "Choice [1]").Trim()
$servers = @{ "" = "www.am.remote.management"; "1" = "www.am.remote.management"; "2" = "www.systemmonitor.us";
              "3" = "www.systemmonitor.co.uk"; "4" = "wwweurope1.systemmonitor.eu.com";
              "5" = "www.system-monitor.com"; "6" = "wwwasia.system-monitor.com" }
if ($servers.ContainsKey($choice)) { $Server = $servers[$choice] } else { $Server = $choice }

# 3. API key -> Windows Credential Manager (never written to a file)
Write-Host "`nPaste the N-sight API key below (input is hidden). It is stored in Windows Credential Manager." -ForegroundColor Cyan
& $VenvPy -m keyring set nsight-mcp nsight_api_key

# 4. Test
$env:NSIGHT_SERVER = $Server
Write-Host "`nTesting connection..." -ForegroundColor Cyan
& $VenvPy "$Dir\server.py" --test
$ok = ($LASTEXITCODE -eq 0)

# 5. Config snippet for Claude desktop
$snippet = [ordered]@{ mcpServers = [ordered]@{ nsight = [ordered]@{
    command = $VenvPy; args = @("$Dir\server.py"); env = [ordered]@{ NSIGHT_SERVER = $Server } } } }
$json = $snippet | ConvertTo-Json -Depth 6
$json | Out-File -Encoding utf8 "$Dir\claude_config_snippet.json"

Write-Host "`n== Add this to Claude desktop ==" -ForegroundColor Cyan
Write-Host "Claude desktop > Settings > Developer > Edit Config, add the 'nsight' entry next to 'halo'"
Write-Host "(also saved to claude_config_snippet.json), then fully quit and reopen Claude.`n"
Write-Host $json
if (-not $ok) { Write-Host "`nThe connection test failed - if it says the server is wrong, re-run and pick another region." -ForegroundColor Yellow }
