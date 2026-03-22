@echo off
cd /d "%~dp0"

if not exist "logs" mkdir "logs"

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HH-mm-ss"') do set DATAHORA=%%i

uv run main.py > "logs\log_pipeline_%DATAHORA%.txt" 2>&1