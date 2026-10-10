# Excel Toolkit - per-user installer (Phase 0 / E6 prototype).
# No administrator rights. Writes only to %LOCALAPPDATA%\ExcelToolkit and HKCU Excel add-in keys.
# Does not change macro security, Trust Center settings or VBA project access.
param(
    [switch]$Rollback,    # switch back to the previously installed version
    [switch]$NoPause
)
$ErrorActionPreference = 'Stop'
$Root = Join-Path $env:LOCALAPPDATA 'ExcelToolkit'
$Versions = Join-Path $Root 'versions'
$StateFile = Join-Path $Root 'current.json'
$LogFile = Join-Path $Root 'install.log'
$OptionsKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
$Package = Split-Path -Parent $MyInvocation.MyCommand.Path

function Write-Log([string]$Message) {
    New-Item -ItemType Directory -Force $Root | Out-Null
    Add-Content -Path $LogFile -Value ("{0} {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)
}

function Stop-Install([string]$Message, [string]$State = 'Nothing was changed.') {
    Write-Log "STOPPED: $Message"
    Write-Host ''
    Write-Host "Excel Toolkit was not installed: $Message" -ForegroundColor Red
    Write-Host $State
    Write-Host "Details: $LogFile"
    if (-not $NoPause) { Read-Host 'Press Enter to close' | Out-Null }
    exit 1
}

function Get-ToolkitOpenValues {
    $props = Get-ItemProperty -Path $OptionsKey -ErrorAction SilentlyContinue
    if (-not $props) { return @() }
    $props.PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' } |
        Select-Object Name, Value
}

function Set-ToolkitAddin([string]$XlamPath) {
    if (-not (Test-Path $OptionsKey)) { New-Item -Path $OptionsKey -Force | Out-Null }
    $open = @(Get-ToolkitOpenValues)
    $existing = $open | Where-Object { $_.Value -like '*\ExcelToolkit\versions\*' } | Select-Object -First 1
    $data = '/R "' + $XlamPath + '"'
    if ($existing) {
        Set-ItemProperty -Path $OptionsKey -Name $existing.Name -Value $data
        return $existing.Name
    }
    $index = 0
    while ($true) {
        $name = if ($index -eq 0) { 'OPEN' } else { "OPEN$index" }
        if (-not ($open | Where-Object { $_.Name -eq $name })) { break }
        $index++
    }
    New-ItemProperty -Path $OptionsKey -Name $name -Value $data -PropertyType String | Out-Null
    return $name
}

function Stop-Worker {
    $pidFile = Join-Path $Root 'state\worker.pid'
    if (Test-Path $pidFile) {
        $workerPid = [int](Get-Content $pidFile -Raw).Trim()
        $process = Get-Process -Id $workerPid -ErrorAction SilentlyContinue
        if ($process -and $process.ProcessName -eq 'ExcelToolkitEngine') {
            Stop-Process -Id $workerPid -Force
            Write-Log "stopped running worker $workerPid"
        }
        Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
    }
    Get-Process -Name 'ExcelToolkitEngine' -ErrorAction SilentlyContinue | Where-Object {
        $_.Path -like "$Versions\*"
    } | ForEach-Object { Stop-Process -Id $_.Id -Force; Write-Log "stopped stray engine $($_.Id)" }
}

trap {
    Stop-Install "unexpected error: $($_.Exception.Message)" 'The installation did not complete. Run it again, or run uninstall.cmd to remove any partial install.'
}

# ---------- Preflight ----------
Write-Log "---- install started from $Package (rollback=$Rollback)"
if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    Stop-Install 'Excel is running. Please close every Excel window (save your work first) and run the installer again.'
}
$platform = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue).Platform
if ($platform -and $platform -ne 'x64') {
    Stop-Install "This build needs 64-bit Excel; this PC has $platform Excel."
}

# ---------- Rollback ----------
if ($Rollback) {
    if (-not (Test-Path $StateFile)) { Stop-Install 'There is no installation to roll back.' }
    $state = Get-Content $StateFile -Raw | ConvertFrom-Json
    if (-not $state.previous) { Stop-Install 'There is no previous version to roll back to.' }
    $previousXlam = Join-Path $Versions "$($state.previous)\ExcelToolkit.xlam"
    if (-not (Test-Path $previousXlam)) { Stop-Install "Previous version $($state.previous) is no longer on this PC." }
    Stop-Worker
    $valueName = Set-ToolkitAddin $previousXlam
    @{ version = $state.previous; previous = $state.version; path = $previousXlam } | ConvertTo-Json | Set-Content $StateFile
    Write-Log "rolled back from $($state.version) to $($state.previous) ($valueName)"
    Write-Host "Excel Toolkit rolled back to version $($state.previous). Start Excel to use it." -ForegroundColor Green
    if (-not $NoPause) { Read-Host 'Press Enter to close' | Out-Null }
    exit 0
}

# ---------- Verify the package ----------
$manifestPath = Join-Path $Package 'manifest.json'
if (-not (Test-Path $manifestPath)) { Stop-Install 'manifest.json is missing - the package is incomplete.' }
$manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
$payload = Join-Path $Package 'payload'
foreach ($entry in $manifest.files) {
    $file = Join-Path $payload $entry.path
    if (-not (Test-Path $file)) { Stop-Install "package file missing: $($entry.path)" }
    $hash = (Get-FileHash -Algorithm SHA256 $file).Hash
    if ($hash -ne $entry.sha256) { Stop-Install "package file damaged or altered: $($entry.path)" }
}
Write-Log "verified $($manifest.files.Count) files for version $($manifest.version)"

# ---------- Install ----------
$dest = Join-Path $Versions $manifest.version
$staging = "$dest.staging"
try {
    Stop-Worker
    if (Test-Path $staging) { Remove-Item $staging -Recurse -Force }
    New-Item -ItemType Directory -Force $staging | Out-Null
    Copy-Item -Path (Join-Path $payload '*') -Destination $staging -Recurse -Force
    # Files are verified against the manifest above; only these copies are unblocked.
    Get-ChildItem $staging -Recurse -File | Unblock-File
    if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
    Rename-Item $staging (Split-Path $dest -Leaf)
} catch {
    if (Test-Path $staging) { Remove-Item $staging -Recurse -Force -ErrorAction SilentlyContinue }
    Stop-Install "copying files failed: $($_.Exception.Message)"
}

$previous = $null
if (Test-Path $StateFile) {
    $old = Get-Content $StateFile -Raw | ConvertFrom-Json
    if ($old.version -ne $manifest.version) { $previous = $old.version }
}
$xlam = Join-Path $dest 'ExcelToolkit.xlam'
try {
    $valueName = Set-ToolkitAddin $xlam
} catch {
    Stop-Install "registering the add-in with Excel failed: $($_.Exception.Message)" 'Files were copied but Excel was not changed. Run the installer again or uninstall.'
}
@{ version = $manifest.version; previous = $previous; path = $xlam } | ConvertTo-Json | Set-Content $StateFile

# Keep only the current and previous versions.
Get-ChildItem $Versions -Directory | Where-Object { $_.Name -ne $manifest.version -and $_.Name -ne $previous } |
    ForEach-Object { Remove-Item $_.FullName -Recurse -Force; Write-Log "removed old version $($_.Name)" }

Write-Log "installed $($manifest.version) at $dest; Excel value $valueName; previous=$previous"
Write-Host "Excel Toolkit $($manifest.version) installed. Start Excel to use it (Ctrl+Shift+J audits the selection)." -ForegroundColor Green
if (-not $NoPause) { Read-Host 'Press Enter to close' | Out-Null }
exit 0
