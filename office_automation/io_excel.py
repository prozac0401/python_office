from __future__ import annotations

from pathlib import Path
from typing import Iterable

import random
import pandas as pd

from .utils import ensure_dir

def read_excel_to_df(path: str | Path, sheet_name: int | str = 0) -> pd.DataFrame:
    """Read an Excel sheet into a DataFrame."""
    return pd.read_excel(Path(path), sheet_name=sheet_name)

def write_df_to_excel(
    df: pd.DataFrame,
    path: str | Path,
    sheet_name: str = "Sheet1",
    index: bool = False,
) -> None:
    """Write DataFrame to an Excel file."""
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
    """Split an Excel sheet into multiple workbooks by a column value."""
    df = read_excel_to_df(input_path, sheet_name=sheet_name)
    if column not in df.columns:
        raise KeyError(f"Column '{column}' not found. Available: {list(df.columns)}")

    out_dir = ensure_dir(output_dir)
    paths: list[Path] = []
    for key, sub in df.groupby(column, dropna=False):
        key_str = "NA" if pd.isna(key) else str(key)
        name = f"{prefix}{key_str}.xlsx" if prefix else f"{key_str}.xlsx"
        out_path = out_dir / name
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
    """Merge multiple Excel files by concatenating rows into a single sheet."""
    frames = []
    for p in input_paths:
        df = read_excel_to_df(p, sheet_name=sheet_name)
        if add_source_column:
            df[source_column_name] = str(Path(p).name)
        frames.append(df)
    merged = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    write_df_to_excel(merged, output_path, sheet_name="Merged")
    return Path(output_path)

def generate_sample_excel(path: str | Path, rows: int = 30) -> Path:
    """Create a sample Excel file for demos/tests."""
    rng = random.Random(42)
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
    write_df_to_excel(df, path, sheet_name="Orders")
    return Path(path)
