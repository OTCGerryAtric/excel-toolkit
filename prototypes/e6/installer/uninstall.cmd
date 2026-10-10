@echo off
rem Runs the per-user uninstaller. -ExecutionPolicy Bypass applies to this one process only.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1" %*
