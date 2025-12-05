# XLSX to Markdown GUI를 exe 파일로 빌드하는 PowerShell 스크립트

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "XLSX to Markdown GUI 빌드 스크립트" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 가상 환경 활성화
& ".\venv\Scripts\Activate.ps1"

# PyInstaller 설치 확인 및 설치
Write-Host "PyInstaller 설치 확인 중..." -ForegroundColor Yellow
pip install pyinstaller --quiet

Write-Host ""
Write-Host "exe 파일 빌드 시작..." -ForegroundColor Yellow
Write-Host ""

# PyInstaller로 빌드
pyinstaller --name="XLSX2Markdown" `
    --windowed `
    --onefile `
    --icon=NONE `
    --add-data "xl2md.py;." `
    --hidden-import=PyQt6.QtCore `
    --hidden-import=PyQt6.QtGui `
    --hidden-import=PyQt6.QtWidgets `
    --hidden-import=pandas `
    --hidden-import=openpyxl `
    --hidden-import=xl2md `
    xl2md_gui.py

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "빌드 완료!" -ForegroundColor Green
Write-Host "exe 파일 위치: dist\XLSX2Markdown.exe" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

Read-Host "계속하려면 Enter를 누르세요"

