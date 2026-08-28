WordToExcel.exe
================

기능
- 실행 시 폴더 선택 창 하나만 표시합니다.
- 선택 폴더 바로 아래의 .doc / .docx / .docm 파일을 이름순으로 처리합니다.
- Word 파일 하나당 Excel 시트 하나를 생성합니다.
- 시트명은 Word 파일명(확장자 제외)을 사용합니다.
- Word COM Content.Copy() -> Excel COM Paste() 방식으로 복붙합니다.
- 결과는 선택 폴더에 Word파일_통합.xlsx 로 저장합니다.
- 동일 파일명이 있으면 Word파일_통합_2.xlsx, _3.xlsx ... 방식으로 회피합니다.
- Excel 시트명 31자 제한, 금지문자, 중복명은 자동 보정합니다.
- 일부 문서 처리 실패 시 같은 폴더에 WordToExcel_error.log 를 남깁니다.

실행 조건
- Windows 10/11 64-bit
- Microsoft Word 설치
- Microsoft Excel 설치

사용법
1. WordToExcel.exe 더블클릭
2. Word 파일들이 있는 폴더 선택
3. 처리가 끝나면 같은 폴더에서 Word파일_통합.xlsx 확인

별도 Python 설치는 필요하지 않습니다.
