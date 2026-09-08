@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "instalar_python.ps1"
pause
