@echo off
REM ============================================================
REM  GuardrailIQ - one-time environment setup (Windows)
REM  Creates a local virtual environment and installs deps.
REM ============================================================
setlocal

echo.
echo ==== GuardrailIQ Setup ====
echo.

REM --- locate Python ---
where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python was not found on PATH.
    echo Please install Python 3.10+ from https://www.python.org/downloads/
    echo and re-run this script.
    pause
    exit /b 1
)

REM --- create virtual environment ---
if not exist ".venv\" (
    echo [1/3] Creating virtual environment .venv ...
    python -m venv .venv
) else (
    echo [1/3] Virtual environment .venv already exists - skipping.
)

REM --- activate ---
echo [2/3] Activating virtual environment ...
call .venv\Scripts\activate.bat

REM --- install dependencies ---
echo [3/3] Installing dependencies from requirements.txt ...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo.
echo ==== Setup complete. ====
echo Run the app with:  run.bat
echo.
pause
endlocal
