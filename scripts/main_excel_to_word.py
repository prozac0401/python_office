from __future__ import annotations

import sys
from pathlib import Path

# Allow `python scripts/...` to import the local package without installation
PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

import argparse

from officekit.render_word import create_word_docs_from_excel


def parse_sheet_arg(value: str) -> int | str:
    """Interpret integer-like values as sheet index, otherwise sheet name."""
    v = value.strip()
    if v and v.lstrip("+-").isdigit():
        return int(v)
    return value


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate one Word doc (DOCX) per Excel row (no COM).")
    parser.add_argument("--input", required=True, help="Input xlsx")
    parser.add_argument(
        "--output-dir",
        default=str(Path("outputs") / "word_docs"),
        help="Output directory for docx files",
    )
    parser.add_argument("--sheet", default=0, type=parse_sheet_arg, help="Sheet name or index")
    parser.add_argument(
        "--template",
        default=str(PROJECT_ROOT / "templates" / "word_template.docx"),
        help="DOCX template for docxtpl (optional).",
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
    print(f"Created {len(out_paths)} Word docs in: {args.output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
