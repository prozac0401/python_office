from __future__ import annotations

import sys
from pathlib import Path

# `python scripts/...` 실행 시, 설치 없이 로컬 패키지를 import 가능하게 한다.
PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

import argparse

from officekit.io_excel import generate_sample_excel


def main() -> int:
    parser = argparse.ArgumentParser(description="샘플 Excel 입력 파일을 생성합니다.")
    parser.add_argument("--out", type=str, default=str(Path("data") / "input.xlsx"), help="출력 xlsx 경로")
    parser.add_argument("--rows", type=int, default=30, help="생성할 행 수")
    parser.add_argument("--ko", action="store_true", help="한글 컬럼/값 형식으로 샘플을 생성합니다.")
    args = parser.parse_args()

    out = generate_sample_excel(args.out, rows=args.rows, korean_mode=args.ko)
    print(f"샘플 Excel 생성 완료: {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
