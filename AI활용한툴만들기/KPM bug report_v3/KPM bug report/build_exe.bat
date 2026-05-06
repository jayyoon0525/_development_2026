@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PY=py -3"
where py >nul 2>&1 || set "PY=python"

echo [KPM Issue Report] Installing runtime + PyInstaller...
%PY% -m pip install -q -r "%~dp0requirements.txt" "pyinstaller>=6.0"
if errorlevel 1 (
  echo pip install failed.
  pause
  exit /b 1
)

echo.
echo [KPM Issue Report] Building dist\KPM_Issue_Report.exe (one file, no console)...
REM Hidden imports: avoid ModuleNotFoundError for openpyxl / Excel in one-file builds.
%PY% -m PyInstaller --noconfirm --clean --windowed --onefile ^
  --name "KPM_Issue_Report" ^
  --collect-all openpyxl ^
  --collect-all pandas ^
  --hidden-import=et_xmlfile ^
  --hidden-import=openpyxl.cell._writer ^
  "%~dp0main.py"

if errorlevel 1 (
  echo Build failed.
  pause
  exit /b 1
)

echo.
echo Done: "%~dp0dist\KPM_Issue_Report.exe"
echo Place general_info.xlsx next to the exe if you use a custom file; otherwise an empty template is created on first run.
pause
