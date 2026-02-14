from __future__ import annotations

import sys
from pathlib import Path

# Allow `python scripts/...` to import the local package without installation
PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

import argparse

from officekit.render_ppt import create_ppt_from_excel


def parse_sheet_arg(value: str) -> int | str:
    """Interpret integer-like values as sheet index, otherwise sheet name."""
    v = value.strip()
    if v and v.lstrip("+-").isdigit():
        return int(v)
    return value


def main() -> int:
    parser = argparse.ArgumentParser(description="Create a PPTX report from an Excel file (data-driven; no COM).")
    parser.add_argument("--input", required=True, help="Input xlsx")
    parser.add_argument("--output", default=str(Path("outputs") / "report.pptx"), help="Output pptx")
    parser.add_argument("--sheet", default=0, type=parse_sheet_arg, help="Sheet name or index")
    parser.add_argument("--title", default="Excel to PPT Report", help="Presentation title")
    parser.add_argument("--group-col", default="Category", help="Column to summarize for chart")
    parser.add_argument("--qty-col", default="Qty")
    parser.add_argument("--unit-col", default="UnitPrice")
    parser.add_argument(
        "--template",
        default=str(PROJECT_ROOT / "templates" / "ppt_template.pptx"),
        help="Base PPTX template (optional).",
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
    print(f"Wrote PPTX: {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
