# Excel Toolkit - per-user uninstaller (Phase 0 / E6 prototype).
# Removes only what the installer added: its Excel OPEN entry, %LOCALAPPDATA%\ExcelToolkit, running engine.
param([switch]$NoPause)
$ErrorActionPreference = 'Stop'
$Root = Join-Path $env:LOCALAPPDATA 'ExcelToolkit'
$OptionsKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
$LogFile = Join-Path $env:TEMP 'ExcelToolkit-uninstall.log'

function Write-Log([string]$Message) {
    Add-Content -Path $LogFile -Value ("{0} {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)
}

Write-Log '---- uninstall started'
if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    Write-Host 'Excel is running. Please close every Excel window and run the uninstaller again. Nothing was changed.' -ForegroundColor Red
    Write-Log 'STOPPED: Excel running'
    if (-not $NoPause) { Read-Host 'Press Enter to close' | Out-Null }
    exit 1
}

Get-Process -Name 'ExcelToolkitEngine' -ErrorAction SilentlyContinue | Where-Object { $_.Path -like "$Root\*" } |
    ForEach-Object { Stop-Process -Id $_.Id -Force; Write-Log "stopped engine $($_.Id)" }

# Remove our OPEN value and renumber the later ones so Excel's list has no gap.
$props = Get-ItemProperty -Path $OptionsKey -ErrorAction SilentlyContinue
if ($props) {
    $open = $props.PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' } |
        Sort-Object { if ($_.Name -eq 'OPEN') { 0 } else { [int]$_.Name.Substring(4) } } |
        Select-Object Name, Value
    $keep = @($open | Where-Object { $_.Value -notlike '*\ExcelToolkit\versions\*' })
    if ($keep.Count -ne @($open).Count) {
        foreach ($entry in $open) { Remove-ItemProperty -Path $OptionsKey -Name $entry.Name }
        for ($i = 0; $i -lt $keep.Count; $i++) {
            $name = if ($i -eq 0) { 'OPEN' } else { "OPEN$i" }
            New-ItemProperty -Path $OptionsKey -Name $name -Value $keep[$i].Value -PropertyType String | Out-Null
        }
        Write-Log "removed Excel Toolkit OPEN entry; kept $($keep.Count) other add-in(s)"
    }
}

if (Test-Path $Root) {
    Remove-Item $Root -Recurse -Force
    Write-Log "removed $Root"
}
Write-Host 'Excel Toolkit has been removed. Your other add-ins and PERSONAL.XLSB were not changed.' -ForegroundColor Green
if (-not $NoPause) { Read-Host 'Press Enter to close' | Out-Null }
exit 0
