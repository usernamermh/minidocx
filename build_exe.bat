@echo off
setlocal EnableExtensions EnableDelayedExpansion

cd /d "%~dp0"

set "PY_CMD="
set "VENV_DIR=.build-venv"
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
set "DIST_EXE=dist\MiniDocxTray.exe"

where py >nul 2>nul
if %errorlevel%==0 (
    py -3.10 -V >nul 2>nul && set "PY_CMD=py -3.10"
)

if not defined PY_CMD (
    where python >nul 2>nul
    if %errorlevel%==0 (
        python -V >nul 2>nul && set "PY_CMD=python"
    )
)

if not defined PY_CMD (
    echo [ERROR] No usable Python was found.
    echo [ERROR] Install Python 3.10+ first, then run this script again.
    pause
    exit /b 1
)

if not exist "%VENV_DIR%\Scripts\activate.bat" (
    echo [INFO] Creating isolated build environment...
    %PY_CMD% -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
)

if not exist "%VENV_PY%" (
    echo [ERROR] Build environment is incomplete: "%VENV_PY%" not found.
    pause
    exit /b 1
)

echo [INFO] Upgrading packaging tools...
"%VENV_PY%" -m pip install --upgrade pip setuptools wheel
if errorlevel 1 (
    echo [ERROR] Failed to upgrade pip/setuptools/wheel.
    pause
    exit /b 1
)

echo [INFO] Installing project dependencies...
"%VENV_PY%" -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Failed to install project dependencies.
    pause
    exit /b 1
)

echo [INFO] Installing build dependencies...
"%VENV_PY%" -m pip install "pyinstaller>=6,<7"
if errorlevel 1 (
    echo [ERROR] Failed to install PyInstaller.
    pause
    exit /b 1
)

echo [INFO] Cleaning previous build output...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [INFO] Building executable...
"%VENV_PY%" -m PyInstaller MiniDocxTray.spec --noconfirm
if errorlevel 1 (
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

if not exist "%DIST_EXE%" (
    echo [ERROR] Build finished but "%DIST_EXE%" was not produced.
    pause
    exit /b 1
)

echo [SUCCESS] Build finished.
echo [SUCCESS] Output: "%cd%\%DIST_EXE%"
pause
