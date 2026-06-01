@echo off
cd /d "%~dp0"

echo ============================================
echo   T11 - Internal Transaction Elimination
echo   V2.0 -- MPAcc 2025
echo ============================================
echo.

echo [Step 1] Checking Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found!
    echo Please install Python from https://www.python.org/downloads/
    echo IMPORTANT: Check "Add Python to PATH" during installation!
    echo.
    pause
    exit /b 1
)
echo [OK] Python is ready
echo.

echo [Step 2] Checking dependencies...
pip install -r requirements.txt -q 2>nul
echo [OK] Dependencies ready
echo.

echo [Step 3] Starting Streamlit...
echo.
echo     --- Open http://localhost:8501 in your browser ---
echo     --- Close this window to stop the server ---
echo.

streamlit run app.py --server.headless true --server.port 8501

pause
