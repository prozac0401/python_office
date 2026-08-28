WordToExcel v2
==============

변경사항
- 진행률 창 추가: 현재 파일명, n/전체, %, 현재 Excel part 표시
- Excel 파일당 최대 30개 Sheet
- 31개 이상이면 part01, part02 ... 로 자동 분리
- 같은 이름이 있으면 실행 단위로 _2, _3 ... 충돌 회피
- Word/Excel COM Application은 작업 전체에서 1회만 생성 후 재사용
- Word COM Copy() -> Excel COM Paste() 방식 유지

출력 예시
- Word 파일 30개 이하:
  Word파일_통합.xlsx
- Word 파일 31개 이상:
  Word파일_통합_part01.xlsx
  Word파일_통합_part02.xlsx
- 기존 파일 충돌 시:
  Word파일_통합_2_part01.xlsx
  Word파일_통합_2_part02.xlsx

병렬 처리 관련
Word Copy()와 Excel Paste()는 Windows 시스템 Clipboard를 공유합니다.
여러 문서를 동시에 Copy/Paste 하면 Clipboard가 덮어써질 수 있으므로 이 핵심 구간은 안전성을 위해 순차 처리합니다.

실행 조건
- Windows 10/11 64-bit
- Microsoft Word 설치
- Microsoft Excel 설치
- Python 설치 불필요
