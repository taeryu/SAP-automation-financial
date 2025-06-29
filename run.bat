@echo off
echo.
echo ================================
echo 재무 자동화 시스템 실행
echo ================================
echo.

REM 가상환경 활성화
if exist "venv\Scripts\activate.bat" (
    echo 🔄 가상환경 활성화 중...
    call venv\Scripts\activate
) else (
    echo ⚠️ 가상환경이 없습니다. install.bat을 먼저 실행해주세요.
    pause
    exit /b 1
)

REM 메인 프로그램 실행
echo 🚀 재무 자동화 시스템 시작...
echo.
python main.py

echo.
echo 👋 프로그램 종료
pause