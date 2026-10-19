@echo off
rem Double-click to sync the published entries and update the website.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\update-site.ps1"
pause
