@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if not errorlevel 1 (
  py -3 "Image to Text.py"
  if not errorlevel 1 exit /b 0
)
where python >nul 2>nul
if not errorlevel 1 (
  python "Image to Text.py"
  exit /b %errorlevel%
)
echo Python was not found. Run setup.bat first.
pause
exit /b 1
