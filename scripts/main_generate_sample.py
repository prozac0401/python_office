from __future__ import annotations

import sys
from pathlib import Path

# Allow `python scripts/...` to import the local package without installation
PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

import argparse

from office_automation.io_excel import generate_sample_excel


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate sample Excel input.")
    parser.add_argument("--out", type=str, default=str(Path("data") / "input.xlsx"), help="Output xlsx path")
    parser.add_argument("--rows", type=int, default=30, help="Number of rows")
    args = parser.parse_args()

    out = generate_sample_excel(args.out, rows=args.rows)
    print(f"Created sample Excel: {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
