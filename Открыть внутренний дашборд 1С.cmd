@echo off
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0prototypes\labor-market-radar-internal\start_internal_dashboard.ps1"
pause
