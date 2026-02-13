from __future__ import annotations

import sys
from pathlib import Path
import glob

# Allow `python scripts/...` to import the local package without installation
PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

import argparse

from office_automation.io_excel import read_excel_to_df, write_df_to_excel, split_excel_by_column, merge_excels_rows
from office_automation.transform import add_total


def cmd_modify(args) -> int:
    df = read_excel_to_df(args.input, sheet_name=args.sheet)
    df2 = add_total(df, qty_col=args.qty_col, unit_price_col=args.unit_col, out_col=args.out_col)
    write_df_to_excel(df2, args.output, sheet_name="Modified")
    print(f"Wrote modified Excel: {args.output}")
    return 0


def cmd_split(args) -> int:
    out_paths = split_excel_by_column(
        args.input,
        args.output_dir,
        column=args.column,
        sheet_name=args.sheet,
        prefix=args.prefix,
    )
    print(f"Split into {len(out_paths)} files:")
    for p in out_paths:
        print(f" - {p}")
    return 0


def _expand_inputs(patterns: list[str]) -> list[Path]:
    """Expand wildcard patterns (Windows-friendly)."""
    expanded: list[str] = []
    for pat in patterns:
        matches = glob.glob(pat)
        if matches:
            expanded.extend(matches)
        else:
            expanded.append(pat)
    # de-dup while preserving order
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
    print(f"Merged {len(inputs)} files into: {out}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Excel operations (no COM): modify / split / merge")
    sub = parser.add_subparsers(dest="cmd", required=True)

    # modify
    p1 = sub.add_parser("modify", help="Add a Total column and write a new Excel file.")
    p1.add_argument("--input", required=True)
    p1.add_argument("--output", required=True)
    p1.add_argument("--sheet", default=0)
    p1.add_argument("--qty-col", default="Qty")
    p1.add_argument("--unit-col", default="UnitPrice")
    p1.add_argument("--out-col", default="Total")
    p1.set_defaults(func=cmd_modify)

    # split
    p2 = sub.add_parser("split", help="Split an Excel file into multiple files by unique values of a column.")
    p2.add_argument("--input", required=True)
    p2.add_argument("--output-dir", required=True)
    p2.add_argument("--column", required=True)
    p2.add_argument("--sheet", default=0)
    p2.add_argument("--prefix", default="")
    p2.set_defaults(func=cmd_split)

    # merge
    p3 = sub.add_parser("merge", help="Merge multiple Excel files by stacking rows.")
    p3.add_argument("--inputs", nargs="+", required=True, help="Input xlsx paths (wildcards like *.xlsx are OK).")
    p3.add_argument("--output", required=True)
    p3.add_argument("--sheet", default=0)
    p3.add_argument("--no-source", action="store_true", help="Do not add a SourceFile column.")
    p3.set_defaults(func=cmd_merge)

    args = parser.parse_args()
    return int(args.func(args))


if __name__ == "__main__":
    raise SystemExit(main())
