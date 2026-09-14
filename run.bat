@echo off
REM ============================================================
REM  GuardrailIQ - launch the Streamlit app (Windows)
REM  Run setup.bat once before using this script.
REM ============================================================
setlocal

REM --- activate venv if present, else use system Python ---
if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
) else (
    echo [WARN] .venv not found - running with system Python.
    echo        Tip: run setup.bat first for an isolated environment.
)

echo.
echo Starting GuardrailIQ ... a browser tab will open at http://localhost:8501
echo Press Ctrl+C in this window to stop the app.
echo.

streamlit run app\guardrailiq_app.py

endlocal
