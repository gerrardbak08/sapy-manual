# 사랑과평안의교회 메뉴얼 제정 · 업무현황 취합 플랫폼

부서/팀별 업무 범위와 현황을 온라인으로 취합하여 교회 규정 메뉴얼을 자동 제정하는 시스템입니다.

## 주요 기능

- **업무현황 취합 폼** (`index.html`) - 12개 부서/팀이 업무 항목을 입력·제출
- **관리자 대시보드** (`admin.html`) - 제출 현황 모니터링 (전체 현황 / 제출 담당자 뷰 토글)
- **Google Sheets 자동 저장** (`google-apps-script.gs`) - 제출 데이터를 3개 시트에 기록
- **메뉴얼 자동 업데이트** (`update_manuals.py`) - 제출 데이터로 `.docx` 파일 갱신 + 수정 이력 테이블 삽입

## 부서 및 팀

**부서 (8):** 예배부, 재정부, 새가족부, 교육부, 친교봉사부, 시설관리부, 환경미화부, 경조사부
**팀 (4):** 문화사역팀, 문서사역팀, 성가대, 찬양팀

## 배포 구조

```
GitHub Pages  →  index.html (폼) + admin.html (대시보드)
Google Apps Script  →  데이터 수신 및 Google Sheets 저장
Python 로컬 실행  →  update_manuals.py로 .docx 파일 업데이트
```

## Google Apps Script 설정

1. [Google Apps Script](https://script.google.com)에서 새 프로젝트 생성
2. `google-apps-script.gs` 내용 붙여넣기
3. 웹 앱으로 배포 (실행 계정: 나, 접근 권한: 모든 사용자)
4. 배포 URL을 `index.html`의 `APPS_SCRIPT_URL` 상수에 입력

## Python 스크립트 사용법

```bash
pip install python-docx requests google-auth google-api-python-client

# 전체 부서 업데이트
python update_manuals.py --all --sheet-id YOUR_SHEET_ID

# 특정 부서만
python update_manuals.py --dept 예배부 --sheet-id YOUR_SHEET_ID

# 미리보기 (파일 수정 없음)
python update_manuals.py --preview --sheet-id YOUR_SHEET_ID
```

## GitHub Pages 활성화

저장소 Settings → Pages → Branch: `main` / `/ (root)` 선택 후 저장
