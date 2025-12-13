# 오프라인 평가 점수 취합하기

## 🚀 설명
- 각 심사위원들이 온라인으로 심사 팀의 점수를 입력하여 제출하면 Google Sheet에 결과 저장
- 심사항목 별 최대점수 지정 가능, 총 합계 점수 자동 계산
- 반응형 html로 웹/모바일 모두 입력 가능

## 🛠️ 기술 스택
- html
- GAS(Google Apps Script)

## 🎯 설치 및 실행 방법
1. 구글 Sheet에서 새 파일 생성한 뒤 결과 취합 시트 명 Scores로 지정(gas 5행)
2. 확장프로그램 - Appscript - Code.gs 에 메모장 파일 내용 복붙
   `AppScript_심사표.txt`
3. 배포 - 새배포 - 유형 선택 - 웹앱에서 엑세스 권한을 모든 사용자로 설정 후 배포
4. 웹앱 url 복사(/exec로 끝남)
5. 구글 sites, github.io 등 웹페이지에 html 게시:
   `index.html`
6. 웹앱 url 복붙(html 229행) 및 웹페이지 실행