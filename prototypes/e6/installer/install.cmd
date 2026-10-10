@echo off
rem Runs the per-user installer. -ExecutionPolicy Bypass applies to this one process only;
rem it does not change the PC's PowerShell policy.
rem Clear any inherited module path (e.g. from PowerShell 7) so Windows PowerShell loads its own modules.
set "PSModulePath="
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %*
