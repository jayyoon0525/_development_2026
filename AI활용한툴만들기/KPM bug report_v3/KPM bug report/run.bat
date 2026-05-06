@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo [IVI 버그 리포트] 실행 중...
echo.

REM 빌드된 exe가 있으면 그걸 실행 (build_exe.bat / build_kpm_fresh.bat 로 생성).
if exist "%~dp0dist\KPM_Issue_Report.exe" (
  echo dist\KPM_Issue_Report.exe 실행...
  start "" "%~dp0dist\KPM_Issue_Report.exe"
  exit /b 0
)

where py >nul 2>&1
if %errorlevel% equ 0 (
  py -3 "%~dp0main.py"
  if errorlevel 1 goto :fail
  exit /b 0
)

where python >nul 2>&1
if %errorlevel% equ 0 (
  python "%~dp0main.py"
  if errorlevel 1 goto :fail
  exit /b 0
)

echo [오류] Python을 찾을 수 없습니다.
echo.
echo  - Python 설치: https://www.python.org/downloads/
echo  - 설치 시 "Add python.exe to PATH" 체크
echo.
goto :pause_end

:fail
echo.
echo 오류로 종료되었습니다. 흔한 해결:
echo   1^) 이 폴더에서 PowerShell을 열고 실행:
echo        python -m pip install -r requirements.txt
echo        python main.py
echo   2^) ModuleNotFoundError 가 나오면 위 1번의 pip 줄이 필요합니다.
echo.

:pause_end
pause
