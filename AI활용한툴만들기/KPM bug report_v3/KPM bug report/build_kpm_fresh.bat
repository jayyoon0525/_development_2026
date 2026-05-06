@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PY=py -3"
where py >nul 2>&1 || set "PY=python"

title KPM Issue Report — clean build

echo ============================================
echo  KPM Issue Report — clean one-file build
echo ============================================
echo.

echo [1/3] Removing previous dist\ and build\ ...
if exist "%~dp0dist" (
  rmdir /s /q "%~dp0dist"
  echo       Removed dist
)
if exist "%~dp0build" (
  rmdir /s /q "%~dp0build"
  echo       Removed build
)
echo.

echo [2/3] pip install (requirements + PyInstaller^) ...
%PY% -m pip install -q -r "%~dp0requirements.txt" "pyinstaller>=6.0"
if errorlevel 1 (
  echo pip install failed.
  pause
  exit /b 1
)
echo       OK
echo.

echo [3/3] PyInstaller --onefile --windowed ...
%PY% -m PyInstaller --noconfirm --clean --windowed --onefile ^
  --name "KPM_Issue_Report" ^
  --collect-all openpyxl ^
  --collect-all pandas ^
  --hidden-import=et_xmlfile ^
  --hidden-import=openpyxl.cell._writer ^
  "%~dp0main.py"

if errorlevel 1 (
  echo.
  echo Build failed.
  pause
  exit /b 1
)

echo.
echo ============================================
echo  Output: "%~dp0dist\KPM_Issue_Report.exe"
echo ============================================
echo  general_info.xlsx 는 exe와 같은 폴더에 두세요.
echo  첫 실행 시 없으면 빈 템플릿이 생성됩니다.
echo ============================================
pause
