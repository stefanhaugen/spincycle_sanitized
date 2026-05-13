@echo off
setlocal EnableDelayedExpansion
title 140SpinCycle - One-Time Setup (requires internet)
color 0E

echo ============================================================
echo   140SpinCycle - ONE-TIME SETUP
echo   Run this ONCE on an internet-connected machine.
echo   After this, the entire folder works offline.
echo ============================================================
echo.

cd /d "%~dp0"
set "APP_DIR=%~dp0"
set "PY_DIR=%APP_DIR%python"
set "PY_ZIP=%TEMP%\python-3.11.9-embed-amd64.zip"
set "PY_URL=https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip"

:: ── Check if already set up ─────────────────────────────────
if exist "%PY_DIR%\python.exe" (
    echo [OK] Python 3.11 is already embedded in this folder.
    echo      You can distribute the folder now.
    echo.
    pause
    exit /b 0
)

:: ─────────────────────────────────────────────────────────────
:: STEP 1: Download Python 3.11 embeddable
:: ─────────────────────────────────────────────────────────────
echo [1/5] Downloading Python 3.11.9 embeddable package...
powershell -ExecutionPolicy Bypass -Command ^
  "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%PY_URL%' -OutFile '%PY_ZIP%' -UseBasicParsing"
if not exist "%PY_ZIP%" (
    echo [ERROR] Download failed. Check your internet connection.
    pause
    exit /b 1
)
echo [OK] Downloaded.

:: ─────────────────────────────────────────────────────────────
:: STEP 2: Extract into python\ folder
:: ─────────────────────────────────────────────────────────────
echo [2/5] Extracting Python into folder...
if not exist "%PY_DIR%" mkdir "%PY_DIR%"
powershell -ExecutionPolicy Bypass -Command ^
  "Expand-Archive -Path '%PY_ZIP%' -DestinationPath '%PY_DIR%' -Force"
if not exist "%PY_DIR%\python.exe" (
    echo [ERROR] Extraction failed.
    pause
    exit /b 1
)
del "%PY_ZIP%" >nul 2>&1
echo [OK] Extracted to python\ folder.

:: ─────────────────────────────────────────────────────────────
:: STEP 3: Fix the ._pth file so pip and packages actually work
::
:: The embeddable Python ships with python311._pth which:
::   - Locks sys.path to ONLY the listed directories
::   - Comments out "import site" so site-packages is invisible
:: We need to: uncomment import site AND add Lib\site-packages
:: ─────────────────────────────────────────────────────────────
echo [3/5] Configuring Python paths...
set "PTH_FILE=%PY_DIR%\python311._pth"
if exist "%PTH_FILE%" (
    :: Rewrite the ._pth file with correct contents
    (
        echo python311.zip
        echo .
        echo Lib\site-packages
        echo import site
    ) > "%PTH_FILE%"
    echo [OK] python311._pth updated.
) else (
    echo [WARN] python311._pth not found — skipping path config.
)

:: Create the Lib\site-packages directory that pip will install into
if not exist "%PY_DIR%\Lib\site-packages" mkdir "%PY_DIR%\Lib\site-packages"

:: ─────────────────────────────────────────────────────────────
:: STEP 4: Bootstrap pip using the bundled wheel
:: ─────────────────────────────────────────────────────────────
echo [4/5] Installing pip...
"%PY_DIR%\python.exe" "%APP_DIR%bootstrap_pip.py"
if %errorlevel% neq 0 (
    echo [ERROR] pip bootstrap failed.
    pause
    exit /b 1
)
echo [OK] pip installed.

:: Verify pip is working
"%PY_DIR%\python.exe" -m pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] pip installed but cannot be found. Path configuration issue.
    pause
    exit /b 1
)
echo [OK] pip verified.

:: ─────────────────────────────────────────────────────────────
:: STEP 5: Install all app packages from local wheels (offline)
:: ─────────────────────────────────────────────────────────────
echo [5/5] Installing app packages from local cache...
"%PY_DIR%\python.exe" -m pip install --no-index --find-links="%APP_DIR%packages" -r "%APP_DIR%requirements.txt"
if %errorlevel% neq 0 (
    echo [ERROR] Package installation failed.
    echo         Make sure the "packages" folder has not been modified.
    pause
    exit /b 1
)
echo [OK] All packages installed.

:: ─────────────────────────────────────────────────────────────
:: Verify streamlit is importable
:: ─────────────────────────────────────────────────────────────
"%PY_DIR%\python.exe" -c "import streamlit; print(f'Streamlit {streamlit.__version__} OK')"
if %errorlevel% neq 0 (
    echo [ERROR] Streamlit failed to import. Something went wrong.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   SETUP COMPLETE!
echo   This folder is now fully portable and works offline.
echo   Distribute it via USB, network share, etc.
echo   Users just double-click "launch_140SpinCycle.bat"
echo ============================================================
echo.
pause
