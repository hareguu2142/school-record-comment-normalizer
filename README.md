# 학생부 세부능력 기록 정리기

학교생활기록부 조회 화면에서 내려받은 인쇄 보고서형 `.xlsx` 파일을 학생별·과목별 표로 정규화하는 Flask 웹앱입니다.

**웹앱:** https://school-record-comment-normalizer.vercel.app

## 기능

- 반복되는 제목, 학급, 열 머리글, 페이지 번호, 학교명 제거
- 생략된 과목·학년·학기 값 복원
- 페이지 경계에서 나뉜 동일 학생 기록 연결
- `과목 / 학년 / 학기 / 번호 / 성명 / 세부능력 및 특기사항` 6개 열로 출력
- 원본 파일 수정 없이 새 `.xlsx` 다운로드

## 개인정보 안내

공개 웹사이트에서 선택한 파일은 변환을 위해 Vercel 서버로 전송됩니다. 애플리케이션은 파일을 데이터베이스나 파일시스템에 영구 저장하지 않지만, 학생 개인정보 처리에 관한 학교·기관의 내부 지침을 확인한 후 이용하세요. 민감한 실제 자료는 가능하면 로컬 실행 방식을 권장합니다.

## 로컬 실행

Windows에서는 `웹앱_실행.bat`를 더블클릭합니다. 처음 실행할 때 `.venv`를 만들고 필요한 패키지를 설치한 다음 `http://127.0.0.1:8765`를 엽니다.

또는 다음 명령으로 실행할 수 있습니다.

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe app.py
```

## Vercel 배포

Vercel CLI 로그인 후 프로젝트 루트에서 실행합니다.

```powershell
vercel --prod
```

Vercel Functions의 요청·응답 본문 제한을 고려해 파일 크기를 4MB로 제한합니다.

## 기술 구성

- Python 3.12
- Flask
- openpyxl
- Vercel Python Runtime

## 라이선스

MIT
