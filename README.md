# OfficeKit (파이썬 전용, Win32 COM 미사용)

Windows 환경에서 **Excel 수정/분할/병합**, **Excel 데이터 기반 PPT 생성**, **Excel 행 단위 Word 문서 생성**을 **Win32 COM 없이** 처리하는 예제 프로젝트입니다.

## 설치
```bash
pip install -r requirements.txt
```

## 템플릿 (기본 제공)
- `templates/excel_template.xlsx`: 입력 형식 참고용
- `templates/word_template.docx`: `{{ 변수 }}` 치환 방식 템플릿
  - 변수명은 Excel 컬럼명과 동일하게 사용합니다. (예: `{{ ID }}`, `{{ Name }}`)
  - 기본 동작은 **python-docx 기반 단순 치환**입니다.
  - 더 복잡한 조건/반복 템플릿이 필요하면 `docxtpl`을 추가 설치해 사용할 수 있습니다.
- `templates/ppt_template.pptx`: 3슬라이드 템플릿
  - 2번 슬라이드: `TABLE_AREA`(사각형) 위치/크기에 표 삽입
  - 3번 슬라이드: `CHART_AREA`(사각형) 위치/크기에 차트 이미지 삽입
  - 템플릿을 수정할 때는 **도형 이름(shape name)을 유지**해야 합니다.

## 한국어 호환 템플릿 세트
아래 명령으로 한국어 컬럼/식별자 템플릿 세트를 생성할 수 있습니다.
```bash
python scripts/main_generate_ko_templates.py
```

생성 파일:
- `templates/excel_template_ko.xlsx` (컬럼명: `아이디`, `이름`, `부서`, `분류`, `수량`, `단가`, `주문일자`, `이메일`)
- `templates/word_template_ko.docx` (치환 변수 예: `{{ 아이디 }}`, `{{ 이름 }}`, `{{ 합계 }}`)
- `templates/ppt_template_ko.pptx` (도형명: `표_영역`, `차트_영역`)

호환 규칙:
- 컬럼명은 영어/한글 별칭을 자동 인식합니다.  
  예: `Qty <-> 수량`, `UnitPrice <-> 단가`, `Total <-> 합계`, `Category <-> 분류`
- PPT 템플릿 식별자는 `TABLE_AREA/CHART_AREA`와 `표_영역/차트_영역`을 모두 지원합니다.

## 0) 샘플 입력 생성
```bash
python scripts/main_generate_sample.py --out data/input.xlsx --rows 30

# 한글 컬럼/값 샘플 생성
python scripts/main_generate_sample.py --out data/input_ko.xlsx --rows 30 --ko
```

## 1) Excel 수정 / 분할 / 병합
```bash
# (1) 수정: Total 컬럼 추가
python scripts/main_excel_ops.py modify --input data/input.xlsx --output outputs/modified.xlsx

# (2) 분할: Department 컬럼 기준
python scripts/main_excel_ops.py split --input data/input.xlsx --output-dir outputs/split_by_dept --column Department

# (2-1) 한글 컬럼 분할: 부서 컬럼 기준
python scripts/main_excel_ops.py split --input data/input_ko.xlsx --output-dir outputs/split_by_dept_ko --column 부서

# (3) 병합: 여러 xlsx를 행 기준으로 합치기
python scripts/main_excel_ops.py merge --inputs outputs/split_by_dept/*.xlsx --output outputs/merged.xlsx
# (PowerShell에서도 OK: 스크립트가 *.xlsx 패턴을 내부에서 확장합니다.)
```

## 2) Excel -> PPT (데이터 기반 생성)
```bash
python scripts/main_excel_to_ppt.py --input data/input.xlsx --output outputs/report.pptx --title "엑셀 보고서"
# 사용자 템플릿 사용:
# python scripts/main_excel_to_ppt.py ... --template path/to/your_template.pptx

# 한국어 템플릿 사용:
python scripts/main_excel_to_ppt.py --input data/input_ko.xlsx --output outputs/report_ko.pptx --template templates/ppt_template_ko.pptx --title "한글 보고서"
```

## 3) Excel 각 행 -> Word 문서 생성
```bash
python scripts/main_excel_to_word.py --input data/input.xlsx --output-dir outputs/word_docs
# 사용자 템플릿 사용:
# python scripts/main_excel_to_word.py ... --template path/to/your_template.docx

# 한국어 템플릿 사용:
python scripts/main_excel_to_word.py --input data/input_ko.xlsx --output-dir outputs/word_docs_ko --template templates/word_template_ko.docx
```

## 참고 / 제약
- COM API를 사용하지 않으므로 Excel의 **서식 체계 전체** 또는 **차트/피벗의 완전한 복제**는 범위 밖입니다.
- 대량 문서 생성은 템플릿 + 데이터 기반 방식으로 사용하는 것을 권장합니다.

