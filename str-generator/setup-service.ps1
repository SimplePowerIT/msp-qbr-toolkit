<#
  QBR STR generator - install on a dedicated PC as a background service account
  ---------------------------------------------------------------------------------------
  Run on the NEW PC, in an elevated PowerShell (Run as Administrator), from the copied halo-mcp folder:

      powershell -ExecutionPolicy Bypass -File .\str-generator\setup-service.ps1

  What it does
    1. Copies halo-mcp to a shared location (default C:\QBR\halo-mcp), skipping .venv, logs and output.
    2. Creates the local service account (default svc-qbr) unless it already exists, or uses a domain account.
    3. Installs Python for all users if needed and builds a fresh .venv.
    4. Opens a window running AS the service account to store the 3 keys in its Credential Manager (hidden input).
    5. Tests Halo, N-sight and SharePoint as the service account.
    6. Registers the scheduled task: starts at boot, runs whether anyone is signed in or not (password saved).
    7. Stops the PC from sleeping or hibernating on AC power.
    8. Starts the task and shows the log.
  Re-enter keys only (no copy / Python rebuild):  ... setup-service.ps1 -KeysOnly
  Afterwards: disable the old task on the previous PC (see the last message the script prints).
  No keys are written to files or shown on screen.
#>
param(
    [string]$InstallDir  = "C:\QBR\halo-mcp",
    [string]$ServiceUser = "svc-qbr",          # local account name, or DOMAIN\user for a domain account
    [switch]$SkipKeys,                          # use if the keys are already stored for this account
    [string]$HaloClientId = "",                 # Halo API application Client ID for this PC (prompted if blank)
    [switch]$KeysOnly                           # re-enter keys + test, then (re)register the task; skips copy and Python rebuild
)

$ErrorActionPreference = "Stop"
$TaskName = "QBR STR Generator"
function Step($t) { Write-Host "`n== $t ==" -ForegroundColor Cyan }
function Ok($t)   { Write-Host "  $t" -ForegroundColor Green }
function Warn($t) { Write-Host "  $t" -ForegroundColor Yellow }
function Fail($t) { Write-Host "  $t" -ForegroundColor Red; exit 1 }

# --- 0. Admin check -------------------------------------------------------------------------
$id = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $id.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Fail "Run this from an elevated PowerShell (right-click > Run as administrator)."
}
$SourceRoot = Split-Path $PSScriptRoot -Parent          # the halo-mcp folder this script came from
Write-Host "QBR STR generator - service install" -ForegroundColor Cyan
Write-Host "  From: $SourceRoot`n  To:   $InstallDir`n  Account: $ServiceUser"

# --- 1. Copy the folder ---------------------------------------------------------------------
Step "1. Copy files"
if ($KeysOnly) {
    Ok "Skipped (-KeysOnly)"
} elseif ((Resolve-Path $SourceRoot).Path.TrimEnd('\') -ieq $InstallDir.TrimEnd('\')) {
    Ok "Already running from $InstallDir"
} else {
    New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
    $rc = (Start-Process robocopy -ArgumentList "`"$SourceRoot`" `"$InstallDir`" /E /NFL /NDL /NJH /NJS /NP /XD .venv __pycache__ output _test /XF *.log" -Wait -PassThru -NoNewWindow).ExitCode
    if ($rc -ge 8) { Fail "Copy failed (robocopy code $rc)." }
    Ok "Copied (without .venv, logs and old output)"
}
$Gen = Join-Path $InstallDir "str-generator"
foreach ($f in @("server.py", "requirements.txt", "claude_config_snippet.json", "nsight-mcp\server.py",
                 "nsight-mcp\claude_config_snippet.json", "str-generator\run.py", "str-generator\config.json",
                 "templates\Client_MonthDATE_Strategic_Technology_Review_Template.docx")) {
    if (-not (Test-Path (Join-Path $InstallDir $f))) { Fail "Missing $f in $InstallDir - copy the whole halo-mcp folder from the old PC." }
}
Get-ChildItem $InstallDir -Recurse -File -Include *.ps1,*.py | Unblock-File -ErrorAction SilentlyContinue

# --- 2. Service account ---------------------------------------------------------------------
Step "2. Service account"
$isDomain = $ServiceUser.Contains("\")
$Account  = if ($isDomain) { $ServiceUser } else { "$env:COMPUTERNAME\$ServiceUser" }
$pw = $null
if (-not $isDomain) {
    $existing = Get-LocalUser -Name $ServiceUser -ErrorAction SilentlyContinue
    if (-not $existing) {
        Write-Host "  Choose a password for the new local account $ServiceUser (long and random; store it in your password manager)."
        $pw  = Read-Host "  Password" -AsSecureString
        $pw2 = Read-Host "  Confirm"  -AsSecureString
        $p1 = [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($pw))
        $p2 = [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($pw2))
        if ($p1 -ne $p2) { Fail "Passwords don't match." }
        $p1 = $null; $p2 = $null
        New-LocalUser -Name $ServiceUser -Password $pw -FullName "QBR generator" `
            -Description "Runs the STR/QBR generator task" -PasswordNeverExpires -UserMayNotChangePassword | Out-Null
        Add-LocalGroupMember -Group "Users" -Member $ServiceUser -ErrorAction SilentlyContinue
        Ok "Created $Account (standard user, not an administrator)"
    } else {
        Ok "$Account already exists"
    }
} else {
    Ok "Using domain account $Account"
}
# Check the password before using it (3 tries); a local account's password can be reset here if forgotten
function Test-Pw([string]$acct, [securestring]$sp) {
    Add-Type -AssemblyName System.DirectoryServices.AccountManagement
    $plainPw = [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($sp))
    try {
        if ($acct.StartsWith("$env:COMPUTERNAME\")) {
            $ctx = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('Machine', $env:COMPUTERNAME)
        } else {
            $ctx = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('Domain', $acct.Split('\')[0])
        }
        return $ctx.ValidateCredentials($acct.Split('\')[-1], $plainPw)
    } catch { return $false } finally { $plainPw = $null }
}
if ($pw -and -not (Test-Pw $Account $pw)) { $pw = $null }
$tries = 0
while (-not $pw) {
    $tries++
    $try = Read-Host "  Password for $Account" -AsSecureString
    if (Test-Pw $Account $try) { $pw = $try; Ok "Password accepted"; break }
    Warn "That password isn't correct for $Account."
    if ($tries -ge 3) {
        if (-not $isDomain) {
            $ans = Read-Host "  Reset the password for local account $ServiceUser now? (y/n)"
            if ($ans -match '^[Yy]') {
                $n1 = Read-Host "  New password" -AsSecureString
                $n2 = Read-Host "  Confirm" -AsSecureString
                $a1 = [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($n1))
                $a2 = [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($n2))
                if ($a1 -ne $a2) { $a1 = $null; $a2 = $null; Fail "Passwords don't match." }
                $a1 = $null; $a2 = $null
                Set-LocalUser -Name $ServiceUser -Password $n1
                Warn "Password reset. Keys stored under the old password can no longer be read, so they will be stored again in step 4."
                if ($SkipKeys) { $SkipKeys = $false }
                $pw = $n1; break
            }
        }
        Fail "Stopping. Re-run when you have the password for $Account."
    }
}
$Cred = New-Object System.Management.Automation.PSCredential($Account, $pw)

# Folder rights: the account needs to write output, logs and the venv cache
& icacls $InstallDir /grant "${Account}:(OI)(CI)M" /T /Q | Out-Null
Ok "Gave $Account modify rights on $InstallDir"

# --- 3. Python for all users + venv ---------------------------------------------------------
Step "3. Python environment"
$Venv = Join-Path $InstallDir ".venv"
$Py  = Join-Path $Venv "Scripts\python.exe"
$PyW = Join-Path $Venv "Scripts\pythonw.exe"
if ($KeysOnly -and (Test-Path $Py)) {
    Ok "Skipped (-KeysOnly); using $Venv"
} else {
function Find-MachinePython {
    $c = @()
    foreach ($root in @($env:ProgramFiles, ${env:ProgramFiles(x86)})) {
        if ($root) { $c += Get-ChildItem -Path $root -Filter "Python3*" -Directory -ErrorAction SilentlyContinue |
                         ForEach-Object { Join-Path $_.FullName "python.exe" } }
    }
    $c | Where-Object { Test-Path $_ } | Sort-Object -Descending | Select-Object -First 1
}
$BasePy = Find-MachinePython
if (-not $BasePy) {
    Warn "No all-users Python found in Program Files (a per-user install can't be used by $ServiceUser)."
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Write-Host "  Installing Python 3.13 for all users with winget..."
        winget install -e --id Python.Python.3.13 --scope machine --silent --accept-package-agreements --accept-source-agreements | Out-Null
        $BasePy = Find-MachinePython
    }
    if (-not $BasePy) { Fail "Install Python 3.10+ for ALL users (python.org installer > Customize > 'Install for all users'), then re-run." }
}
$ver = & $BasePy -c "import sys;print('%d.%d'%sys.version_info[:2])"
if ([version]$ver -lt [version]"3.10") { Fail "Python $ver at $BasePy is too old; 3.10+ is required." }
Ok "Python $ver at $BasePy"
$Venv = Join-Path $InstallDir ".venv"
if (Test-Path $Venv) { Remove-Item $Venv -Recurse -Force }
& $BasePy -m venv $Venv
$Py  = Join-Path $Venv "Scripts\python.exe"
$PyW = Join-Path $Venv "Scripts\pythonw.exe"
& $Py -m pip install --quiet --upgrade pip
& $Py -m pip install --quiet -r (Join-Path $InstallDir "requirements.txt")
$nreq = Join-Path $InstallDir "nsight-mcp\requirements.txt"
if (Test-Path $nreq) { & $Py -m pip install --quiet -r $nreq }
& icacls $Venv /grant "${Account}:(OI)(CI)M" /T /Q | Out-Null
Ok "Built $Venv"
}

# --- 4. Keys, stored as the service account ---------------------------------------------------
Step "4. Halo API application and keys for $Account"
# Halo Client ID: saved in str-generator\config.json (env.HALO_CLIENT_ID), which overrides claude_config_snippet.json
$cfgPath = Join-Path $Gen "config.json"
$cfgObj  = Get-Content $cfgPath -Raw | ConvertFrom-Json
$snipObj = Get-Content (Join-Path $InstallDir "claude_config_snippet.json") -Raw | ConvertFrom-Json
$curId = $null
if ($cfgObj.PSObject.Properties.Name -contains "env" -and $cfgObj.env.HALO_CLIENT_ID) { $curId = $cfgObj.env.HALO_CLIENT_ID }
if (-not $curId) { $curId = $snipObj.mcpServers.halo.env.HALO_CLIENT_ID }
if (-not $HaloClientId) {
    Write-Host "  Halo API application Client ID currently used: $curId"
    $HaloClientId = (Read-Host "  Paste the Client ID of the Halo app this PC should use (Enter keeps the current one)").Trim()
}
if ($HaloClientId -and $HaloClientId -ne $curId) {
    if (-not ($cfgObj.PSObject.Properties.Name -contains "env")) { $cfgObj | Add-Member -NotePropertyName env -NotePropertyValue ([pscustomobject]@{}) }
    $cfgObj.env | Add-Member -NotePropertyName HALO_CLIENT_ID -NotePropertyValue $HaloClientId -Force
    $cfgObj | ConvertTo-Json -Depth 6 | Out-File -Encoding utf8 $cfgPath
    Ok "Halo Client ID set to $HaloClientId (saved in str-generator\config.json)"
    Warn "Paste THIS application's client secret at 1/3 in the key window."
} else {
    Ok "Halo Client ID: $curId"
}
$helper = Join-Path $InstallDir "_store_keys.ps1"
@'
param([string]$Py)
$Host.UI.RawUI.WindowTitle = "Store STR generator keys as $env:USERNAME"
Write-Host "This window runs as $env:USERNAME. Paste each key, then press Enter (input is hidden)." -ForegroundColor Cyan
Write-Host "Tip: right-click pastes in this window. Press Enter on an empty line to keep a key that is already stored." -ForegroundColor Cyan
$keys = @(
    @{ svc = "halo-mcp";      name = "halo_client_secret";  label = "Halo API client secret" },
    @{ svc = "nsight-mcp";    name = "nsight_api_key";      label = "N-sight API key" },
    @{ svc = "str-generator"; name = "graph_client_secret"; label = "SharePoint app (Entra) client secret VALUE" }
)
$i = 0
foreach ($k in $keys) {
    $i++
    $env:STR_KSVC = $k.svc; $env:STR_KNAME = $k.name
    $len = & $Py -c "import os,keyring;v=keyring.get_password(os.environ['STR_KSVC'],os.environ['STR_KNAME']);print(len(v or ''))"
    $has = ([int]$len) -gt 0
    while ($true) {
        Write-Host ""
        Write-Host ("{0}/3  {1}{2}" -f $i, $k.label, $(if ($has) { "  (already stored - Enter keeps it)" } else { "" })) -ForegroundColor Cyan
        $sec = Read-Host "Paste" -AsSecureString
        $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec))
        $plain = $plain.Trim()
        if (-not $plain) {
            if ($has) { Write-Host "  Kept the stored value." -ForegroundColor Green; break }
            Write-Host "  Nothing was pasted - try again (right-click to paste)." -ForegroundColor Yellow
            continue
        }
        $env:STR_KVAL = $plain
        & $Py -c "import os,keyring;keyring.set_password(os.environ['STR_KSVC'],os.environ['STR_KNAME'],os.environ['STR_KVAL'])"
        $env:STR_KVAL = $null; $plain = $null
        $len = & $Py -c "import os,keyring;v=keyring.get_password(os.environ['STR_KSVC'],os.environ['STR_KNAME']);print(len(v or ''))"
        if (([int]$len) -gt 0) { Write-Host ("  Saved ({0} characters)." -f $len) -ForegroundColor Green; break }
        Write-Host "  Saving failed - try again." -ForegroundColor Yellow
    }
}
$env:STR_KSVC = $null; $env:STR_KNAME = $null
Write-Host ""
Write-Host "All three keys stored. This window will close." -ForegroundColor Green
Start-Sleep 3
'@ | Out-File -Encoding utf8 $helper
& icacls $helper /grant "${Account}:R" /Q | Out-Null
if ($SkipKeys) {
    Ok "Skipped (-SkipKeys)"
} else {
    Write-Host "  A new window will open running as $Account. Paste the three keys there."
    Start-Process powershell.exe -Credential $Cred -LoadUserProfile -WorkingDirectory $InstallDir `
        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$helper`" -Py `"$Py`"" -Wait
    Ok "Keys stored in $Account's Credential Manager"
}
Remove-Item $helper -Force -ErrorAction SilentlyContinue

# --- 5. Test as the service account ---------------------------------------------------------
Step "5. Test connections as $Account"
$testPy  = Join-Path $InstallDir "_svc_test.py"
$testOut = Join-Path $InstallDir "_svc_test.txt"
@'
import sys, traceback
sys.path.insert(0, r"STRDIR")
import str_generator as G
cfg = G.load_config()
def chk(label, fn):
    try:
        fn(); print("OK   " + label)
    except Exception as e:
        print("FAIL " + label + ": " + str(e)[:200])
chk("Halo", lambda: G.halo_mod().test_connection())
chk("N-sight", lambda: G.nsight_mod().test_connection())
def sp():
    from sharepoint import SharePoint
    s = SharePoint(cfg)
    if not s.list_id("STR Requests"):
        raise RuntimeError("STR Requests list not found")
chk("SharePoint", sp)
'@.Replace("STRDIR", $Gen) | Out-File -Encoding utf8 $testPy
Start-Process $Py -Credential $Cred -LoadUserProfile -WorkingDirectory $Gen -ArgumentList "`"$testPy`"" `
    -RedirectStandardOutput $testOut -RedirectStandardError "$testOut.err" -Wait -WindowStyle Hidden
$res = Get-Content $testOut -ErrorAction SilentlyContinue
$res | ForEach-Object { if ($_ -like "OK*") { Ok $_ } else { Warn $_ } }
$err = Get-Content "$testOut.err" -ErrorAction SilentlyContinue | Where-Object { $_ -match "Error|Traceback" }
Remove-Item $testPy, $testOut, "$testOut.err" -Force -ErrorAction SilentlyContinue
if (-not $res -or ($res | Where-Object { $_ -like "FAIL*" })) {
    if ($err) { $err | Select-Object -Last 3 | ForEach-Object { Warn $_ } }
    Fail "Fix the failed connection(s) above (usually a mistyped key: re-run with the same options), then re-run this script."
}

# --- 6. Scheduled task ----------------------------------------------------------------------
Step "6. Scheduled task"
# The account needs "Log on as a batch job" to run while nobody is signed in (error 2147943785 without it)
$sid  = (New-Object System.Security.Principal.NTAccount($Account)).Translate([System.Security.Principal.SecurityIdentifier]).Value
$inf  = Join-Path $env:TEMP "str_batchright.inf"; $sdb = Join-Path $env:TEMP "str_batchright.sdb"
secedit /export /cfg $inf /areas USER_RIGHTS | Out-Null
$lines = Get-Content $inf
$batch = $lines | Where-Object { $_ -match '^SeBatchLogonRight' }
if ($batch -and $batch -match [regex]::Escape("*$sid")) {
    Ok "$Account already has 'Log on as a batch job'"
} else {
    if ($batch) { $lines = $lines -replace '^(SeBatchLogonRight\s*=\s*.*)$', "`$1,*$sid" }
    else        { $lines = $lines -replace '^\[Privilege Rights\]$', "[Privilege Rights]`r`nSeBatchLogonRight = *$sid" }
    $lines | Set-Content $inf -Encoding Unicode
    secedit /configure /db $sdb /cfg $inf /areas USER_RIGHTS | Out-Null
    gpupdate /force | Out-Null
    Ok "Granted $Account 'Log on as a batch job'"
}
Remove-Item $inf, $sdb -ErrorAction SilentlyContinue
$plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($pw))
$action   = New-ScheduledTaskAction -Execute $PyW -Argument "`"$Gen\run.py`" --poll" -WorkingDirectory $Gen
$trigger  = New-ScheduledTaskTrigger -AtStartup
$trigger.Delay = "PT1M"                                  # give the network a minute after boot
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable `
            -ExecutionTimeLimit (New-TimeSpan -Days 0) -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 5) `
            -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
    -User $Account -Password $plain -RunLevel Limited -Force | Out-Null
$plain = $null
Ok "Registered '$TaskName': at startup, runs whether $ServiceUser is signed in or not (password saved)"

# --- 7. Power settings ----------------------------------------------------------------------
Step "7. Keep the PC awake"
powercfg /change standby-timeout-ac 0
powercfg /change hibernate-timeout-ac 0
Ok "Sleep and hibernate disabled on AC power"

# --- 8. Start and check ---------------------------------------------------------------------
Step "8. Start"
Start-ScheduledTask -TaskName $TaskName
Write-Host "  Waiting 40 seconds for the first poll..."
Start-Sleep 40
$log = Join-Path $Gen "str-generator.log"
if (Test-Path $log) { Get-Content $log -Tail 5 | ForEach-Object { Write-Host "  $_" } }
$state = (Get-ScheduledTask -TaskName $TaskName).State
Ok "Task state: $state"

Write-Host "`nDone." -ForegroundColor Green
Write-Host "LAST STEP - on the OLD PC (your laptop), in PowerShell, turn off the old task so only one machine answers requests:" -ForegroundColor Yellow
Write-Host "    Disable-ScheduledTask -TaskName `"$TaskName`"" -ForegroundColor Yellow
Write-Host "Then submit a test request from the Generate STR button. Log: $log"
