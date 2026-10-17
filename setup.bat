@echo off
setlocal
cd /d "%~dp0"
echo === Vendor Balance Confirmation - setup ===

set "PY="
py -3 --version >nul 2>nul && set "PY=py -3"
if not defined PY python --version >nul 2>nul && set "PY=python"
if not defined PY if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"

if not defined PY (
    echo Python not found. Installing Python 3.12 for this user via winget...
    winget install -e --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements
    if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
)
if not defined PY (
    echo.
    echo Could not find or install Python. Install Python 3.10+ from https://www.python.org/downloads/
    echo then run setup.bat again.
    pause
    exit /b 1
)

echo Using Python: %PY%
if not exist ".venv\Scripts\python.exe" (
    %PY% -m venv .venv || goto :fail
)
".venv\Scripts\python.exe" -m pip install --upgrade pip || goto :fail
".venv\Scripts\python.exe" -m pip install -r requirements.txt || goto :fail

echo.
echo Setup complete. Start the app with run.bat
pause
exit /b 0

:fail
echo.
echo Setup failed - see the messages above.
pause
exit /b 1
