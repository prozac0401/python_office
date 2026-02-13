from __future__ import annotations

import pandas as pd

def add_total(df: pd.DataFrame, qty_col: str = "Qty", unit_price_col: str = "UnitPrice", out_col: str = "Total") -> pd.DataFrame:
    """Add a Total column = Qty * UnitPrice."""
    for col in (qty_col, unit_price_col):
        if col not in df.columns:
            raise KeyError(f"Missing column '{col}'. Available: {list(df.columns)}")
    out = df.copy()
    out[out_col] = pd.to_numeric(out[qty_col], errors="coerce").fillna(0) * pd.to_numeric(out[unit_price_col], errors="coerce").fillna(0)
    return out

def summarize_by_group(df: pd.DataFrame, group_col: str, value_col: str, agg: str = "sum") -> pd.DataFrame:
    """Group and aggregate a value column."""
    if group_col not in df.columns:
        raise KeyError(f"Missing group column '{group_col}'.")
    if value_col not in df.columns:
        raise KeyError(f"Missing value column '{value_col}'.")
    if agg not in ("sum", "mean", "count", "max", "min"):
        raise ValueError("agg must be one of: sum, mean, count, max, min")
    grouped = getattr(df.groupby(group_col, dropna=False)[value_col], agg)().reset_index()
    grouped = grouped.sort_values(by=value_col, ascending=False)
    return grouped
