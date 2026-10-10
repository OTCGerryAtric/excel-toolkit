@echo off
rem Runs the per-user installer. -ExecutionPolicy Bypass applies to this one process only;
rem it does not change the PC's PowerShell policy.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" %*
