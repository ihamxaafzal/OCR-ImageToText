@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Luma OCR Setup

echo ========================================
echo       Luma OCR - Automatic Setup
echo ========================================
echo.

set "PYEXE="
where py >nul 2>nul
if not errorlevel 1 (
  py -3 -c "import sys; raise SystemExit(0 if sys.version_info >= (3,9) else 1)" >nul 2>nul
  if not errorlevel 1 set "PYEXE=py -3"
)
if not defined PYEXE (
  where python >nul 2>nul
  if not errorlevel 1 (
    python -c "import sys; raise SystemExit(0 if sys.version_info >= (3,9) else 1)" >nul 2>nul
    if not errorlevel 1 set "PYEXE=python"
  )
)

if not defined PYEXE (
  echo Python 3.9 or newer was not found. Trying to install Python 3.12 with winget...
  where winget >nul 2>nul
  if errorlevel 1 goto no_python
  winget install --id Python.Python.3.12 -e --scope user --silent --accept-package-agreements --accept-source-agreements
  if errorlevel 1 goto python_install_failed
  if exist "%LocalAppData%\Programs\Python\Python312\python.exe" set "PYEXE="%LocalAppData%\Programs\Python\Python312\python.exe""
  if not defined PYEXE (
    where py >nul 2>nul
    if not errorlevel 1 set "PYEXE=py -3.12"
  )
  if not defined PYEXE goto python_install_failed
)

echo.
echo Installing Python dependencies from requirements.txt...
%PYEXE% -m pip install --upgrade --no-warn-script-location -r requirements.txt
if errorlevel 1 goto pip_failed

echo.
where tesseract >nul 2>nul
if not errorlevel 1 goto ready
if exist "%ProgramFiles%\Tesseract-OCR\tesseract.exe" goto ready
if exist "%LocalAppData%\Programs\Tesseract-OCR\tesseract.exe" goto ready

echo Tesseract OCR engine was not found. Trying to install it with winget...
where winget >nul 2>nul
if errorlevel 1 goto no_winget
winget install --id UB-Mannheim.TesseractOCR -e --silent --accept-package-agreements --accept-source-agreements
if errorlevel 1 goto tesseract_install_failed
where tesseract >nul 2>nul
if not errorlevel 1 goto ready
if exist "%ProgramFiles%\Tesseract-OCR\tesseract.exe" goto ready
if exist "%LocalAppData%\Programs\Tesseract-OCR\tesseract.exe" goto ready
goto tesseract_install_failed

:ready
echo.
echo Setup complete. Python packages and Tesseract OCR are ready.
echo Double-click run.bat to launch the app.
echo.
echo If you need Urdu OCR, install the Urdu (urd) trained data for Tesseract first.
pause
exit /b 0

:no_python
echo ERROR: Python was not found and winget is unavailable.
echo Install Python 3.9+ from https://www.python.org/downloads/ and run setup.bat again.
goto failed

:python_install_failed
echo ERROR: Could not install or locate Python. Install Python 3.9+ and run setup.bat again.
goto failed

:pip_failed
echo ERROR: Python package installation failed. Check your internet connection and try again.
goto failed

:no_winget
echo ERROR: winget is not available, so Tesseract could not be installed automatically.
echo Install Tesseract OCR, then run setup.bat again.
goto failed

:tesseract_install_failed
echo ERROR: Tesseract was not installed or could not be confirmed.
echo Rerun setup.bat and approve the Tesseract installer, or install Tesseract manually.
goto failed

:failed
echo.
pause
exit /b 1

