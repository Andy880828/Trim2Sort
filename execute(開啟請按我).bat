@echo off
cd /d "%~dp0"

REM 清除可能存在的舊虛擬環境變數，避免與其他專案衝突
set VIRTUAL_ENV=
set VENV_DIR=

echo Checking if uv is installed...
where uv >nul 2>nul

IF %ERRORLEVEL% NEQ 0 (
    echo uv not found. Installing uv...
    powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
)

echo Adding module...
uv sync

echo Running program with uv...
uv run python -m src.main

pause
