@echo off
setlocal

cd /d "%~dp0"

set "PY_CMD="
set "USE_VENV=1"
set "VENV_DIR=D:\codes\venvs\docs\.venv"

where python >nul 2>nul
if %errorlevel%==0 (
    python -V >nul 2>nul
    if %errorlevel%==0 set "PY_CMD=python"
)

if not defined PY_CMD (
    where py >nul 2>nul
    if %errorlevel%==0 (
        py -V >nul 2>nul
        if %errorlevel%==0 set "PY_CMD=py"
    )
)

if not defined PY_CMD (
    echo [ERROR] No working Python found in PATH.
    echo [TIP] Please make sure python can run in cmd, or reinstall Python and add it to PATH.
    pause
    exit /b 1
)

if not exist "%VENV_DIR%" (
    echo [INFO] Creating virtual environment...
    if not exist "D:\codes\venvs\docs" mkdir "D:\codes\venvs\docs"
    %PY_CMD% -m venv "%VENV_DIR%"
)

if not exist "%VENV_DIR%\Scripts\activate.bat" (
    echo [WARN] Virtual environment creation failed. Falling back to current Python.
    set "USE_VENV=0"
)

if "%USE_VENV%"=="1" (
    call "%VENV_DIR%\Scripts\activate.bat"
    set "RUN_PY=python"
) else (
    set "RUN_PY=%PY_CMD%"
)

%RUN_PY% -c "import http.server, webbrowser" >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Your current Python is missing required standard modules.
    pause
    exit /b 1
)

echo [INFO] Trying direct launch...
%RUN_PY% server.py
if errorlevel 1 (
    echo [ERROR] Program exited with an error.
    echo [TIP] If this happened right after you pressed Ctrl+C, you can ignore it.
    pause
    exit /b 1
)
