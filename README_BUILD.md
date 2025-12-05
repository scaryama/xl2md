# XLSX to Markdown GUI - EXE 빌드 가이드

## 개요
Python GUI 애플리케이션을 Windows 실행 파일(.exe)로 변환하는 방법입니다.

## 사전 준비

### 1. 필요한 패키지 설치
```bash
# 가상 환경 활성화
venv\Scripts\activate

# PyInstaller 설치
pip install pyinstaller
```

## 빌드 방법

### 방법 1: 배치 파일 사용 (권장)
```bash
build_exe.bat
```

### 방법 2: PowerShell 스크립트 사용
```powershell
.\build_exe.ps1
```

### 방법 3: 수동 빌드
```bash
# 가상 환경 활성화
venv\Scripts\activate

# PyInstaller 실행
pyinstaller --name="XLSX2Markdown" ^
    --windowed ^
    --onefile ^
    --add-data "xl2md.py;." ^
    --hidden-import=PyQt6.QtCore ^
    --hidden-import=PyQt6.QtGui ^
    --hidden-import=PyQt6.QtWidgets ^
    --hidden-import=pandas ^
    --hidden-import=openpyxl ^
    --hidden-import=xl2md ^
    xl2md_gui.py
```

## 빌드 옵션 설명

- `--name="XLSX2Markdown"`: 생성될 exe 파일 이름
- `--windowed`: 콘솔 창 없이 GUI만 표시
- `--onefile`: 단일 exe 파일로 생성 (폴더 대신)
- `--add-data "xl2md.py;."`: xl2md.py 모듈을 포함
- `--hidden-import=...`: 필요한 모듈들을 명시적으로 포함

## 빌드 결과

빌드가 완료되면 다음 위치에 exe 파일이 생성됩니다:
```
dist\XLSX2Markdown.exe
```

## 배포

`dist\XLSX2Markdown.exe` 파일만 복사하여 다른 Windows 컴퓨터에서 실행할 수 있습니다.
Python이나 추가 라이브러리 설치가 필요 없습니다.

## 문제 해결

### PyQt6 모듈을 찾을 수 없는 경우
```bash
pyinstaller --hidden-import=PyQt6.QtCore --hidden-import=PyQt6.QtGui --hidden-import=PyQt6.QtWidgets ...
```

### pandas 모듈 오류
```bash
pyinstaller --hidden-import=pandas --hidden-import=openpyxl ...
```

### 파일 크기가 큰 경우
`--onefile` 옵션을 제거하면 여러 파일로 생성되지만 크기가 줄어듭니다:
```bash
pyinstaller --windowed --name="XLSX2Markdown" xl2md_gui.py
```

## 참고 사항

- 첫 빌드는 시간이 걸릴 수 있습니다 (몇 분 소요)
- exe 파일 크기는 약 50-100MB 정도입니다 (모든 의존성 포함)
- Windows Defender나 백신 프로그램이 경고할 수 있지만, 이는 정상입니다 (서명되지 않은 실행 파일)

