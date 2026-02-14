from __future__ import annotations

import sys
from pathlib import Path

# `python scripts/...` 실행 시, 설치 없이 로컬 패키지를 import 가능하게 한다.
PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

import argparse

from officekit.render_ppt import create_ppt_from_excel


def parse_sheet_arg(value: str) -> int | str:
    """정수처럼 보이면 시트 인덱스로, 아니면 시트 이름으로 해석한다."""
    v = value.strip()
    if v and v.lstrip("+-").isdigit():
        return int(v)
    return value


def main() -> int:
    parser = argparse.ArgumentParser(description="Excel 파일에서 데이터 기반 PPTX 보고서를 생성합니다(COM 미사용).")
    parser.add_argument("--input", required=True, help="입력 xlsx 경로")
    parser.add_argument("--output", default=str(Path("outputs") / "report.pptx"), help="출력 pptx 경로")
    parser.add_argument("--sheet", default=0, type=parse_sheet_arg, help="시트 이름 또는 인덱스")
    parser.add_argument("--title", default="엑셀 보고서", help="프레젠테이션 제목")
    parser.add_argument("--group-col", default="Category", help="차트 요약 기준 컬럼(예: Category 또는 분류)")
    parser.add_argument("--qty-col", default="Qty", help="수량 컬럼명(예: Qty 또는 수량)")
    parser.add_argument("--unit-col", default="UnitPrice", help="단가 컬럼명(예: UnitPrice 또는 단가)")
    parser.add_argument(
        "--template",
        default=str(PROJECT_ROOT / "templates" / "ppt_template.pptx"),
        help="기반 PPTX 템플릿(선택, TABLE_AREA/CHART_AREA 또는 표_영역/차트_영역 지원)",
    )
    args = parser.parse_args()

    tpl = Path(args.template)
    template_pptx = str(tpl) if tpl.exists() else None

    out = create_ppt_from_excel(
        args.input,
        args.output,
        sheet_name=args.sheet,
        title=args.title,
        group_col=args.group_col,
        qty_col=args.qty_col,
        unit_price_col=args.unit_col,
        template_pptx=template_pptx,
    )
    print(f"PPTX 생성 완료: {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
