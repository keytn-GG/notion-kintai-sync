@echo off
rem notion-kintai-sync: manual run (menu)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0src\run.ps1" -Mode Menu
pause
