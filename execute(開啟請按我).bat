@echo off
cd /d "%~dp0"

REM Clear old virtual environment variables to avoid conflicts with other projects
set VIRTUAL_ENV=
set VENV_DIR=

REM Check network connection
echo Checking network connection...
ping -n 1 8.8.8.8 >nul 2>nul
IF %ERRORLEVEL% NEQ 0 (
    echo Warning: Network connection check failed. Git pull may not work.
    echo Continuing anyway...
) ELSE (
    echo Network connection OK.
)

REM Check if Git is installed
echo Checking if Git is installed...
where git >nul 2>nul
IF %ERRORLEVEL% NEQ 0 (
    echo Git not found. Attempting to install Git...
    where winget >nul 2>nul
    IF %ERRORLEVEL% EQU 0 (
        echo Installing Git via winget...
        winget install --id Git.Git -e --source winget --accept-package-agreements --accept-source-agreements
        REM Reload PATH environment variable
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

REM Execute Git Pull
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

REM Set micromamba path and R environment path
set MAMBA=%~dp0dependencies\micromamba.exe
set R_ENV=%~dp0.renv

REM Check if R environment exists, create if not exists
if not exist "%R_ENV%" (
    echo Creating R environment at %R_ENV%...
    "%MAMBA%" create -y -p "%R_ENV%" ^
        -c conda-forge ^
        r-base ^
        rpy2
    IF %ERRORLEVEL% NEQ 0 (
        echo Warning: Failed to create R environment. Continuing anyway...
    ) ELSE (
        echo R environment created successfully.
    )
) ELSE (
    echo R environment already exists at %R_ENV%.
)

echo Running program with uv...
uv run python -m src.main

pause
