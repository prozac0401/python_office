from __future__ import annotations

import sys
from pathlib import Path

# `python scripts/...` 실행 시, 설치 없이 로컬 패키지를 import 가능하게 한다.
PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

import argparse

from officekit.render_word import create_word_docs_from_excel


def parse_sheet_arg(value: str) -> int | str:
    """정수처럼 보이면 시트 인덱스로, 아니면 시트 이름으로 해석한다."""
    v = value.strip()
    if v and v.lstrip("+-").isdigit():
        return int(v)
    return value


def main() -> int:
    parser = argparse.ArgumentParser(description="Excel 각 행마다 Word 문서(DOCX) 1개를 생성합니다(COM 미사용).")
    parser.add_argument("--input", required=True, help="입력 xlsx 경로")
    parser.add_argument(
        "--output-dir",
        default=str(Path("outputs") / "word_docs"),
        help="DOCX 출력 디렉터리",
    )
    parser.add_argument("--sheet", default=0, type=parse_sheet_arg, help="시트 이름 또는 인덱스")
    parser.add_argument(
        "--template",
        default=str(PROJECT_ROOT / "templates" / "word_template.docx"),
        help="docxtpl용 DOCX 템플릿(선택, 영어/한글 컬럼 변수 모두 호환)",
    )
    args = parser.parse_args()

    tpl = Path(args.template)
    template_docx = str(tpl) if tpl.exists() else None

    out_paths = create_word_docs_from_excel(
        args.input,
        args.output_dir,
        sheet_name=args.sheet,
        template_docx=template_docx,
    )
    print(f"Word 문서 {len(out_paths)}개 생성 완료: {args.output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
