@echo off
chcp 65001 > nul
echo =====================================================
echo   Ansys License Tool - EXE 빌드 스크립트
echo =====================================================
echo.

:: 가상환경 활성화 (있다면)
IF EXIST ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
    echo [OK] 가상환경 활성화됨
)

:: 의존성 설치
echo [1/3] 패키지 설치 중...
pip install streamlit pandas reportlab python-pptx pyinstaller --quiet
IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] 패키지 설치 실패
    pause
    exit /b 1
)
echo [OK] 패키지 설치 완료

:: 기존 빌드 폴더 정리
echo [2/3] 빌드 폴더 정리 중...
IF EXIST dist\AnsysLicenseTool rmdir /s /q dist\AnsysLicenseTool
IF EXIST build\AnsysLicenseTool rmdir /s /q build\AnsysLicenseTool
echo [OK] 정리 완료

:: PyInstaller 빌드
echo [3/3] EXE 빌드 중... (시간이 걸릴 수 있습니다)
pyinstaller AnsysLicenseTool.spec --clean
IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] 빌드 실패
    pause
    exit /b 1
)

echo.
echo =====================================================
echo   빌드 완료!
echo   실행파일: dist\AnsysLicenseTool\AnsysLicenseTool.exe
echo =====================================================
echo.

:: 빌드된 EXE 폴더 열기
explorer dist\AnsysLicenseTool
pause
