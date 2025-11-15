@echo off
cd /d "%~dp0"

echo Checking if uv is installed...
where uv >nul 2>nul

IF %ERRORLEVEL% NEQ 0 (
    echo uv not found. Installing uv...
    powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
)

echo Adding module...
uv sync

echo Running program with uv...
uv run main.py

pause
