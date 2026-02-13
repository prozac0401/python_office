# Office Automation (Python only, no Win32 COM)

Windows 환경에서 **Excel 수정/분할/병합**, **Excel 데이터 기반 PPT 생성**, **Excel 행 단위 Word 문서 생성**을
**Win32 COM 없이** 처리하는 예제 프로젝트입니다.

## 설치
```bash
pip install -r requirements.txt
```

## 템플릿(기본 제공)
- `templates/excel_template.xlsx` : 입력 형식 참고용
- `templates/word_template.docx` : `{{ 변수 }}` 치환 방식 템플릿 (변수명 = 엑셀 컬럼명, 예: `{{ ID }}`, `{{ Name }}`)
  - 기본은 **python-docx 기반 치환**으로 동작합니다.
  - 더 복잡한 템플릿(조건/반복 등)이 필요하면 `docxtpl`을 추가 설치해서 사용할 수 있습니다(선택).
- `templates/ppt_template.pptx` : 3슬라이드 템플릿  
  - 2번 슬라이드에 `TABLE_AREA`(사각형) → 테이블이 그 위치/크기로 삽입  
  - 3번 슬라이드에 `CHART_AREA`(사각형) → 차트 이미지가 그 위치/크기로 삽입  
  - 디자인을 바꾸고 싶으면 PowerPoint로 열어 편집하되 **shape name은 유지**하세요.

## 0) 샘플 엑셀 생성
```bash
python scripts/main_generate_sample.py --out data/input.xlsx --rows 30
```

## 1) Excel 수정 / 분할 / 병합
```bash
# (1) 수정: Total 컬럼 추가
python scripts/main_excel_ops.py modify --input data/input.xlsx --output outputs/modified.xlsx

# (2) 분할: Department 컬럼 기준
python scripts/main_excel_ops.py split --input data/input.xlsx --output-dir outputs/split_by_dept --column Department

# (3) 병합: 여러 엑셀을 행으로 합치기
python scripts/main_excel_ops.py merge --inputs outputs/split_by_dept/*.xlsx --output outputs/merged.xlsx
# (PowerShell에서도 OK: 스크립트가 *.xlsx 패턴을 내부에서 확장합니다)
```

## 2) Excel -> PPT (데이터 기반 생성)
```bash
python scripts/main_excel_to_ppt.py --input data/input.xlsx --output outputs/report.pptx --title "Orders Report"
# 커스텀 템플릿 사용:
# python scripts/main_excel_to_ppt.py ... --template path/to/your_template.pptx
```

## 3) Excel 각 행 -> Word 문서로 저장
```bash
python scripts/main_excel_to_word.py --input data/input.xlsx --output-dir outputs/word_docs
# 커스텀 템플릿 사용:
# python scripts/main_excel_to_word.py ... --template path/to/your_template.docx
```

## 참고/제약
- COM을 쓰지 않기 때문에, Excel의 **수식 재계산**이나 **차트/피벗의 완벽한 복제**는 범위 밖입니다.
- 대신 “데이터 기반”으로 PPT/Word를 생성하는 방식을 권장합니다.
