from __future__ import annotations

import pandas as pd

from .column_alias import require_existing_column, resolve_existing_column


def add_total(df: pd.DataFrame, qty_col: str = "Qty", unit_price_col: str = "UnitPrice", out_col: str = "Total") -> pd.DataFrame:
    """`Total = Qty * UnitPrice` 계산 열을 추가한다."""
    qty_name = require_existing_column(df.columns, qty_col, role="수량 열")
    unit_name = require_existing_column(df.columns, unit_price_col, role="단가 열")
    out_name = resolve_existing_column(df.columns, out_col)
    if out_name is None:
        if out_col == "Total" and (qty_name == "수량" or unit_name == "단가"):
            out_name = "합계"
        elif out_col == "합계" and (qty_name == "Qty" or unit_name == "UnitPrice"):
            out_name = "Total"
        else:
            out_name = out_col

    out = df.copy()
    out[out_name] = (
        pd.to_numeric(out[qty_name], errors="coerce").fillna(0)
        * pd.to_numeric(out[unit_name], errors="coerce").fillna(0)
    )
    return out


def summarize_by_group(df: pd.DataFrame, group_col: str, value_col: str, agg: str = "sum") -> pd.DataFrame:
    """지정한 그룹 열 기준으로 값 열을 집계한다."""
    group_name = require_existing_column(df.columns, group_col, role="그룹 열")
    value_name = require_existing_column(df.columns, value_col, role="값 열")
    if agg not in ("sum", "mean", "count", "max", "min"):
        raise ValueError("agg는 sum, mean, count, max, min 중 하나여야 합니다.")
    grouped = getattr(df.groupby(group_name, dropna=False)[value_name], agg)().reset_index()
    grouped = grouped.sort_values(by=value_name, ascending=False)
    return grouped
