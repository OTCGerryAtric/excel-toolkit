@echo off
rem Runs the per-user uninstaller. -ExecutionPolicy Bypass applies to this one process only.
rem Clear any inherited module path (e.g. from PowerShell 7) so Windows PowerShell loads its own modules.
set "PSModulePath="
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1" %*
