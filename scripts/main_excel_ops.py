from __future__ import annotations

import sys
from pathlib import Path
import glob

# `python scripts/...` 실행 시, 설치 없이 로컬 패키지를 import 가능하게 한다.
PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

import argparse

from officekit.io_excel import read_excel_to_df, write_df_to_excel, split_excel_by_column, merge_excels_rows
from officekit.transform import add_total


def parse_sheet_arg(value: str) -> int | str:
    """정수처럼 보이면 시트 인덱스로, 아니면 시트 이름으로 해석한다."""
    v = value.strip()
    if v and v.lstrip("+-").isdigit():
        return int(v)
    return value


def cmd_modify(args) -> int:
    df = read_excel_to_df(args.input, sheet_name=args.sheet)
    df2 = add_total(df, qty_col=args.qty_col, unit_price_col=args.unit_col, out_col=args.out_col)
    write_df_to_excel(df2, args.output, sheet_name="수정결과")
    print(f"수정된 Excel 저장 완료: {args.output}")
    return 0


def cmd_split(args) -> int:
    out_paths = split_excel_by_column(
        args.input,
        args.output_dir,
        column=args.column,
        sheet_name=args.sheet,
        prefix=args.prefix,
    )
    print(f"{len(out_paths)}개 파일로 분할 완료:")
    for p in out_paths:
        print(f" - {p}")
    return 0


def _expand_inputs(patterns: list[str]) -> list[Path]:
    """와일드카드 패턴을 실제 파일 목록으로 확장한다(Windows 친화)."""
    expanded: list[str] = []
    for pat in patterns:
        matches = glob.glob(pat)
        if matches:
            expanded.extend(matches)
        else:
            expanded.append(pat)
    # 순서를 유지하면서 중복 제거
    seen = set()
    out: list[Path] = []
    for s in expanded:
        p = str(Path(s))
        if p not in seen:
            seen.add(p)
            out.append(Path(s))
    return out


def cmd_merge(args) -> int:
    inputs = _expand_inputs(args.inputs)
    out = merge_excels_rows(inputs, args.output, sheet_name=args.sheet, add_source_column=not args.no_source)
    print(f"{len(inputs)}개 파일 병합 완료: {out}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Excel 작업(COM 미사용): modify / split / merge")
    sub = parser.add_subparsers(dest="cmd", required=True)

    # 수정
    p1 = sub.add_parser("modify", help="Total 열을 추가해 새 Excel 파일로 저장합니다.")
    p1.add_argument("--input", required=True)
    p1.add_argument("--output", required=True)
    p1.add_argument("--sheet", default=0, type=parse_sheet_arg)
    p1.add_argument("--qty-col", default="Qty", help="수량 컬럼명(예: Qty 또는 수량)")
    p1.add_argument("--unit-col", default="UnitPrice", help="단가 컬럼명(예: UnitPrice 또는 단가)")
    p1.add_argument("--out-col", default="Total", help="결과 컬럼명(예: Total 또는 합계)")
    p1.set_defaults(func=cmd_modify)

    # 분할
    p2 = sub.add_parser("split", help="지정 열의 고유값 기준으로 Excel 파일을 분할합니다.")
    p2.add_argument("--input", required=True)
    p2.add_argument("--output-dir", required=True)
    p2.add_argument("--column", required=True, help="분할 기준 컬럼명(예: Department 또는 부서)")
    p2.add_argument("--sheet", default=0, type=parse_sheet_arg)
    p2.add_argument("--prefix", default="")
    p2.set_defaults(func=cmd_split)

    # 병합
    p3 = sub.add_parser("merge", help="여러 Excel 파일의 행을 쌓아 하나로 병합합니다.")
    p3.add_argument("--inputs", nargs="+", required=True, help="입력 xlsx 경로들(예: *.xlsx 와일드카드 가능)")
    p3.add_argument("--output", required=True)
    p3.add_argument("--sheet", default=0, type=parse_sheet_arg)
    p3.add_argument("--no-source", action="store_true", help="SourceFile 열을 추가하지 않습니다.")
    p3.set_defaults(func=cmd_merge)

    args = parser.parse_args()
    return int(args.func(args))


if __name__ == "__main__":
    raise SystemExit(main())
