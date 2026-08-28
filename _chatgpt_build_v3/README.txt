WordToExcel v3
================

기능
- 시작 화면에서 Word 파일 폴더를 선택합니다.
- Excel 파일당 최대 Sheet 수를 1~9999 사이에서 직접 지정할 수 있습니다.
- 기본값은 30입니다.
- 입력 Word 파일(.doc/.docx/.docm) 하나당 Excel Sheet 하나를 생성합니다.
- Sheet 이름은 Word 파일명(확장자 제외)을 사용합니다.
- Word COM Content.Copy() -> Excel COM Paste() 방식으로 실제 Clipboard 복붙을 수행합니다.
- 지정한 Sheet 수를 초과하면 다음 Excel 파일로 자동 분리합니다.
- 예: 최대 30 설정 + 65개 성공 -> part01(30), part02(30), part03(5)
- 결과 파일명이 이미 있으면 _2, _3 ... 식으로 충돌을 회피합니다.
- 진행창에 현재 파일명, 진행 개수, 퍼센트, 현재 Excel part를 표시합니다.
- 진행창에 Clipboard 복붙 사용 주의 Alert를 계속 표시합니다.
- 일부 Word 파일 처리 실패 시 나머지는 계속 진행하며 WordToExcel_error.log를 남깁니다.

복붙 주의
- 이 프로그램은 Windows 시스템 Clipboard를 사용합니다.
- 변환 중 Ctrl+C, Ctrl+X, 파일 복사, 캡처/클립보드 유틸리티 등 Clipboard를 사용하는 작업은 피하는 것을 권장합니다.
- 처리 중 Clipboard 내용이 바뀌면 Excel Sheet에 잘못 붙여넣어질 수 있습니다.
- 이 이유로 Copy -> Paste 구간은 병렬화하지 않고 순차 처리합니다.

실행 조건
- Windows 10/11 64-bit
- Microsoft Word 설치
- Microsoft Excel 설치
- Python 별도 설치 불필요(EXE 사용 시)
