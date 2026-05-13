@echo off
setlocal EnableDelayedExpansion
title spincycle Data Pipeline
color 0A

echo ============================================================
echo SPINCYCLE_data program (MassHunter / Chemstation)
echo ============================================================
echo.

cd /d "%~dp0"
set "APP_DIR=%~dp0"
set "PY_DIR=%APP_DIR%python"
set "PYTHON=%PY_DIR%\python.exe"

:: ── Verify embedded Python exists ────────────────────────────
if not exist "%PYTHON%" (
    echo [ERROR] Embedded Python not found!
    echo.
    echo         The "python" folder is missing or incomplete.
    echo         An admin needs to run SETUP_once.bat first
    echo         on an internet-connected machine.
    echo.
    pause
    exit /b 1
)

:: ── Verify the ._pth file is configured ──────────────────────
:: If someone copies the folder before SETUP ran fully, packages
:: won't be on the path. Check and fix if needed.
set "PTH_FILE=%PY_DIR%\python311._pth"
if exist "%PTH_FILE%" (
    findstr /C:"Lib\site-packages" "%PTH_FILE%" >nul 2>&1
    if !errorlevel! neq 0 (
        echo [FIX] Repairing Python path configuration...
        (
            echo python311.zip
            echo .
            echo Lib\site-packages
            echo import site
        ) > "%PTH_FILE%"
    )
)

:: ── Verify packages are installed ────────────────────────────
"%PYTHON%" -c "import streamlit" >nul 2>&1
if %errorlevel% neq 0 (
    echo [SETUP] Packages not found. Installing from local cache...
    if not exist "%PY_DIR%\Lib\site-packages" mkdir "%PY_DIR%\Lib\site-packages"
    :: Bootstrap pip first if needed
    "%PYTHON%" -m pip --version >nul 2>&1
    if !errorlevel! neq 0 (
        echo [SETUP] Bootstrapping pip...
        "%PYTHON%" "%APP_DIR%bootstrap_pip.py"
        if !errorlevel! neq 0 (
            echo [ERROR] pip bootstrap failed. Run SETUP_once.bat again.
            pause
            exit /b 1
        )
    )
    "%PYTHON%" -m pip install --no-index --find-links="%APP_DIR%packages" -r "%APP_DIR%requirements.txt" --quiet
    if !errorlevel! neq 0 (
        echo [ERROR] Package installation failed.
        echo         Try deleting the "python" folder and
        echo         running SETUP_once.bat again.
        pause
        exit /b 1
    )
    echo [OK] Packages installed.
)

echo [OK] Python 3.11.9 (embedded) ready.
echo [OK] All packages verified.
echo.
echo ============================================================
echo   Launching 140SpinCycle in your default browser...
echo   Close this window or press Ctrl+C to stop the server.
echo ============================================================
echo.

"%PYTHON%" -m streamlit run "%APP_DIR%app.py" --server.address=localhost --browser.gatherUsageStats=false

pause
