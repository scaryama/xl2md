@echo off
REM XLSX to Markdown GUI를 exe 파일로 빌드하는 스크립트

echo ========================================
echo XLSX to Markdown GUI 빌드 스크립트
echo ========================================
echo.

REM 가상 환경 활성화
call venv\Scripts\activate.bat

REM PyInstaller 설치 확인 및 설치
echo PyInstaller 설치 확인 중...
pip install pyinstaller --quiet

echo.
echo exe 파일 빌드 시작...
echo.

REM PyInstaller로 빌드
pyinstaller --name="XLSX2Markdown" ^
    --windowed ^
    --onefile ^
    --icon=NONE ^
    --add-data "xl2md.py;." ^
    --hidden-import=PyQt6.QtCore ^
    --hidden-import=PyQt6.QtGui ^
    --hidden-import=PyQt6.QtWidgets ^
    --hidden-import=pandas ^
    --hidden-import=openpyxl ^
    --hidden-import=xl2md ^
    xl2md_gui.py

echo.
echo ========================================
echo 빌드 완료!
echo exe 파일 위치: dist\XLSX2Markdown.exe
echo ========================================
echo.

pause

