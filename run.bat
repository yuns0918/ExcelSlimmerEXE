@echo off
setlocal ENABLEDELAYEDEXPANSION

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

echo =============================================
echo ExcelSuite - Run from source (venv)
echo =============================================

if not exist ".venv_suite" (
  echo [ERROR] .venv_suite not found.
  echo 먼저 install.bat 을 실행해서 가상환경과 패키지를 설치하세요.
  pause
  exit /b 1
)

echo [*] Activating virtual environment...
call ".venv_suite\Scripts\activate.bat"
if errorlevel 1 (
  echo [ERROR] Failed to activate virtual environment.
  pause
  exit /b 1
)

echo [*] Launching ExcelSlimmer (Qt UI)...
python "excel_slimmer_qt.py"

echo.
echo [INFO] Program finished.
pause
