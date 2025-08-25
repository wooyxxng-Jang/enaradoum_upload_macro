# e나라도움 업로드 자동화 프로그램 v1.0

정부 e나라도움 사이트에서 **PDF 파일 자동 업로드 & 검증**을 지원하는 파이썬 GUI 프로그램입니다.

## ✨ 기능

* GUI에서 PDF 폴더 경로 설정 및 저장
* \[F1] 키: 크롬 디버그 모드 브라우저에 연결
* \[F3] 키: PDF 업로드 자동화 시작/중지
* \[ESC] 키: 프로그램 종료
* 업로드 검증 기능 (엑셀/CSV 결과 저장)

## 📦 설치 방법

1. Python 3.10+ 설치
2. 저장소 클론:

   ```bash
   git clone https://github.com/wooyxxng/enaradoum_upload_macro.git
   cd enaradoum_upload_macro
   ```
3. 가상환경 생성 및 패키지 설치:

   ```bash
   python -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   ```

## 🚀 실행 방법

1. 크롬을 **디버그 모드**로 실행:

   ```bash
   chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\chrome_debug"
   ```
   OR 
   ```bash
   run_debug_chrome.bat
   ```
   배치파일 실행

2. 프로그램 실행:

   ```bash
   python main.py
   ```

## 📁 출력 결과

* 업로드 검증 결과: `upload_verification_YYYYMMDD_HHMMSS.xlsx` (또는 `.csv`)

## ⚠️ 주의

* 크롬 브라우저 버전과 `webdriver-manager`에서 설치하는 크롬드라이버 버전이 맞아야 합니다.
