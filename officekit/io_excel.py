from __future__ import annotations

from pathlib import Path
from typing import Iterable

import random
import pandas as pd

from .column_alias import require_existing_column
from .utils import ensure_dir, safe_filename


def _next_available_path(path: Path) -> Path:
    """파일명이 충돌하면 `_N`을 붙여 사용 가능한 경로를 반환한다."""
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    idx = 2
    while True:
        candidate = path.with_name(f"{stem}_{idx}{suffix}")
        if not candidate.exists():
            return candidate
        idx += 1


def read_excel_to_df(path: str | Path, sheet_name: int | str = 0) -> pd.DataFrame:
    """Excel 시트를 DataFrame으로 읽는다."""
    return pd.read_excel(Path(path), sheet_name=sheet_name)


def write_df_to_excel(
    df: pd.DataFrame,
    path: str | Path,
    sheet_name: str = "Sheet1",
    index: bool = False,
) -> None:
    """DataFrame을 Excel 파일로 저장한다."""
    out = Path(path)
    out.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=index)


def split_excel_by_column(
    input_path: str | Path,
    output_dir: str | Path,
    column: str,
    sheet_name: int | str = 0,
    prefix: str = "",
) -> list[Path]:
    """지정한 열의 값 기준으로 Excel을 여러 파일로 분할한다."""
    df = read_excel_to_df(input_path, sheet_name=sheet_name)
    split_col = require_existing_column(df.columns, column, role="분할 기준 열")

    out_dir = ensure_dir(output_dir)
    paths: list[Path] = []
    for key, sub in df.groupby(split_col, dropna=False):
        key_str = "NA" if pd.isna(key) else str(key)
        base_name = f"{prefix}{key_str}" if prefix else key_str
        name = safe_filename(base_name) + ".xlsx"
        out_path = _next_available_path(out_dir / name)
        write_df_to_excel(sub, out_path, sheet_name="Data")
        paths.append(out_path)
    return paths


def merge_excels_rows(
    input_paths: Iterable[str | Path],
    output_path: str | Path,
    sheet_name: int | str = 0,
    add_source_column: bool = True,
    source_column_name: str = "_source",
) -> Path:
    """여러 Excel 파일의 행을 이어 붙여 단일 시트로 병합한다."""
    frames = []
    for p in input_paths:
        df = read_excel_to_df(p, sheet_name=sheet_name)
        if add_source_column:
            df[source_column_name] = str(Path(p).name)
        frames.append(df)
    merged = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    write_df_to_excel(merged, output_path, sheet_name="Merged")
    return Path(output_path)


def generate_sample_excel(path: str | Path, rows: int = 30, korean_mode: bool = False) -> Path:
    """데모/테스트용 샘플 Excel 파일을 생성한다."""
    rng = random.Random(42)
    if korean_mode:
        names = ["지수", "민호", "소라", "현우", "유나", "서준", "하나", "준"]
        depts = ["영업", "마케팅", "재무", "연구개발", "운영"]
        cats = ["가", "나", "다"]
    else:
        names = ["Jisoo", "Minho", "Sora", "Hyunwoo", "Yuna", "Seojun", "Hana", "Jun"]
        depts = ["Sales", "Marketing", "Finance", "R&D", "Operations"]
        cats = ["A", "B", "C"]
    today = pd.Timestamp.today().normalize()

    data = []
    for i in range(1, rows + 1):
        name = rng.choice(names)
        dept = rng.choice(depts)
        cat = rng.choice(cats)
        qty = rng.randint(1, 20)
        price = rng.choice([5_000, 12_000, 20_000, 35_000, 50_000])
        email = f"{name.lower()}{i}@example.com"
        data.append({
            "ID": i,
            "Name": name,
            "Department": dept,
            "Category": cat,
            "Qty": qty,
            "UnitPrice": price,
            "OrderDate": (today - pd.Timedelta(days=rng.randint(0, 14))).date(),
            "Email": email,
        })
    df = pd.DataFrame(data)
    if korean_mode:
        df = df.rename(columns={
            "ID": "아이디",
            "Name": "이름",
            "Department": "부서",
            "Category": "분류",
            "Qty": "수량",
            "UnitPrice": "단가",
            "OrderDate": "주문일자",
            "Email": "이메일",
        })
    write_df_to_excel(df, path, sheet_name="주문데이터" if korean_mode else "Orders")
    return Path(path)
