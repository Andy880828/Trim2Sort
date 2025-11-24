@echo off
cd /d "%~dp0"

REM 清除可能存在的舊虛擬環境變數，避免與其他專案衝突
set VIRTUAL_ENV=
set VENV_DIR=

REM 檢查網路連線
echo Checking network connection...
ping -n 1 8.8.8.8 >nul 2>nul
IF %ERRORLEVEL% NEQ 0 (
    echo Warning: Network connection check failed. Git pull may not work.
    echo Continuing anyway...
) ELSE (
    echo Network connection OK.
)

REM 檢查 Git 是否已安裝
echo Checking if Git is installed...
where git >nul 2>nul
IF %ERRORLEVEL% NEQ 0 (
    echo Git not found. Attempting to install Git...
    where winget >nul 2>nul
    IF %ERRORLEVEL% EQU 0 (
        echo Installing Git via winget...
        winget install --id Git.Git -e --source winget --accept-package-agreements --accept-source-agreements
        REM 重新載入 PATH 環境變數
        call refreshenv >nul 2>nul || (
            echo Please restart the script after Git installation, or add Git to PATH manually.
            echo Git installation may require restarting the command prompt.
        )
    ) ELSE (
        echo winget not found. Please install Git manually from https://git-scm.com/download/win
        echo Continuing without Git pull...
        goto :skip_git_pull
    )
)

REM 執行 Git Pull
echo Updating repository with git pull...
git pull
IF %ERRORLEVEL% NEQ 0 (
    echo Warning: Git pull failed. This may be normal if not in a git repository or network issues.
)

:skip_git_pull

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
