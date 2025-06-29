@echo off
echo.
echo ====================================
echo 재무 자동화 시스템 설치 스크립트
echo ====================================
echo.

echo [1/3] Python 가상환경 생성 중...
python -m venv venv
if %errorlevel% neq 0 (
    echo ❌ 가상환경 생성 실패!
    pause
    exit /b 1
)

echo [2/3] 가상환경 활성화 중...
call venv\Scripts\activate

echo [3/3] 필요한 패키지 설치 중...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ❌ 패키지 설치 실패!
    pause
    exit /b 1
)

echo.
echo ✅ 설치 완료!
echo.
echo 🚀 사용법:
echo   1. venv\Scripts\activate  (가상환경 활성화)
echo   2. python main.py         (프로그램 실행)
echo.
pause