@echo off
setlocal ENABLEDELAYEDEXPANSION

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

echo =============================================
echo ExcelSlimmer - Build single EXE (PyInstaller)
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

echo [*] Cleaning previous build artifacts...
for %%D in (build dist __pycache__) do (
  if exist "%%D" rmdir /s /q "%%D"
)

set "NAME=ExcelSlimmer"

rem 우선 현재 폴더에 있는 ExcelSlimmer.ico 를 사용하고,
rem 없으면 기존 ExcelCleaner\icon.ico 를 사용합니다.
set "ICON=ExcelSlimmer.ico"
if not exist "%ICON%" set "ICON=..\ExcelCleaner\icon.ico"

echo [*] Building single EXE with PyInstaller...
pyinstaller ^
  --onefile ^
  --noconsole ^
  --name %NAME% ^
  --paths "..\ExcelCleaner" ^
  --paths "..\ExcelImageOptimization" ^
  --paths "..\ExcelByteReduce" ^
  --icon "%ICON%" ^
  "excel_suite_pipeline.py"

if errorlevel 1 (
  echo [ERROR] PyInstaller build failed.
  pause
  exit /b 1
)

if not exist "dist\%NAME%.exe" (
  echo [ERROR] EXE not found: dist\%NAME%.exe
  pause
  exit /b 1
)

echo.
echo [OK] Build complete.
echo EXE: "%CD%\dist\%NAME%.exe"
echo 이 파일 하나만 복사해서 다른 PC에 배포하면 됩니다.
echo.
pause
