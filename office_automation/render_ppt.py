from __future__ import annotations

from pathlib import Path

import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

import matplotlib.pyplot as plt

from .io_excel import read_excel_to_df
from .transform import add_total, summarize_by_group
from .utils import ensure_dir


def _find_shape_by_name(slide, name: str):
    for shape in slide.shapes:
        if getattr(shape, "name", "") == name:
            return shape
    return None


def _remove_shape(shape) -> None:
    el = shape._element
    el.getparent().remove(el)


def _set_title(slide, title: str) -> None:
    if slide.shapes.title:
        slide.shapes.title.text = title
        return
    # fallback: first textbox
    for shape in slide.shapes:
        if shape.has_text_frame:
            shape.text_frame.text = title
            return


def _set_subtitle(slide, subtitle: str) -> None:
    # common: placeholder[1] is subtitle on title slide
    try:
        if len(slide.placeholders) > 1:
            slide.placeholders[1].text = subtitle
            return
    except Exception:
        pass
    # fallback: named shape
    shp = _find_shape_by_name(slide, "SUBTITLE")
    if shp and shp.has_text_frame:
        shp.text_frame.text = subtitle


def _get_or_create_slide(prs: Presentation, index: int, layout_index: int):
    if len(prs.slides) > index:
        return prs.slides[index]
    return prs.slides.add_slide(prs.slide_layouts[layout_index])


def _get_region_from_named_shape(slide, shape_name: str, default_region):
    shp = _find_shape_by_name(slide, shape_name)
    if shp is None:
        return default_region, None
    region = (shp.left, shp.top, shp.width, shp.height)
    return region, shp


def _insert_table(slide, df: pd.DataFrame, region, max_rows: int = 12) -> None:
    left, top, width, height = region
    view = df.head(max_rows)

    rows = len(view) + 1
    cols = len(view.columns)

    table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table

    # header
    for j, col in enumerate(view.columns):
        cell = table.cell(0, j)
        cell.text = str(col)
        for p in cell.text_frame.paragraphs:
            p.font.bold = True
            p.font.size = Pt(12)
            p.alignment = PP_ALIGN.CENTER

    # body
    for i in range(len(view)):
        for j, col in enumerate(view.columns):
            val = view.iloc[i, j]
            table.cell(i + 1, j).text = "" if pd.isna(val) else str(val)

    # sizing
    for r in range(rows):
        for c in range(cols):
            tf = table.cell(r, c).text_frame
            for p in tf.paragraphs:
                if p.font.size is None:
                    p.font.size = Pt(11)
                p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER


def _make_bar_chart_png(summary_df: pd.DataFrame, category_col: str, value_col: str, out_path: Path, title: str) -> Path:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fig = plt.figure(figsize=(10, 5.5))
    ax = fig.add_subplot(111)

    x = summary_df[category_col].astype(str).tolist()
    y = summary_df[value_col].tolist()
    ax.bar(x, y)
    ax.set_title(title)
    ax.set_xlabel(category_col)
    ax.set_ylabel(value_col)
    ax.tick_params(axis="x", rotation=30)

    fig.tight_layout()
    fig.savefig(out_path, dpi=200)
    plt.close(fig)
    return out_path


def _insert_picture(slide, img_path: Path, region) -> None:
    left, top, width, height = region
    slide.shapes.add_picture(str(img_path), left, top, width=width, height=height)


def create_ppt_from_excel(
    input_xlsx: str | Path,
    output_pptx: str | Path,
    sheet_name: int | str = 0,
    title: str = "Excel to PPT Report",
    group_col: str = "Category",
    qty_col: str = "Qty",
    unit_price_col: str = "UnitPrice",
    template_pptx: str | Path | None = None,
) -> Path:
    """Create a PPT report from an Excel file (data-driven; no COM).

    If `template_pptx` is provided, it will be used as the base deck.
    The default template shipped with this project expects:
      - Slide 1: Title slide
      - Slide 2: Table slide, with a rectangle shape named "TABLE_AREA"
      - Slide 3: Chart slide, with a rectangle shape named "CHART_AREA"

    If those slides/shapes are missing, the function falls back to a simple layout.
    """
    df = read_excel_to_df(input_xlsx, sheet_name=sheet_name)
    df2 = add_total(df, qty_col=qty_col, unit_price_col=unit_price_col, out_col="Total")
    summary = summarize_by_group(df2, group_col=group_col, value_col="Total", agg="sum")

    # Load template if given
    if template_pptx:
        tpl_path = Path(template_pptx)
        if not tpl_path.exists():
            raise FileNotFoundError(f"Template not found: {tpl_path}")
        prs = Presentation(str(tpl_path))
    else:
        prs = Presentation()

    # ---- Slide 1 (title)
    cover = _get_or_create_slide(prs, 0, 0)
    _set_title(cover, title)
    _set_subtitle(cover, f"Source: {Path(input_xlsx).name}")

    # ---- Slide 2 (table)
    table_slide = _get_or_create_slide(prs, 1, 5)
    _set_title(table_slide, "Preview (top rows)")
    default_table_region = (Inches(0.6), Inches(1.5), Inches(12.2), Inches(5.2))
    region, placeholder = _get_region_from_named_shape(table_slide, "TABLE_AREA", default_table_region)
    if placeholder is not None:
        _remove_shape(placeholder)
    _insert_table(table_slide, df2, region, max_rows=12)

    # ---- Slide 3 (chart)
    chart_slide = _get_or_create_slide(prs, 2, 5)
    _set_title(chart_slide, "Total by Category")
    tmp_dir = ensure_dir(Path(output_pptx).parent / "_tmp")
    chart_png = _make_bar_chart_png(
        summary,
        category_col=group_col,
        value_col="Total",
        out_path=tmp_dir / "total_by_category.png",
        title="Total by Category",
    )
    default_chart_region = (Inches(0.8), Inches(1.4), Inches(11.8), Inches(5.4))
    region2, placeholder2 = _get_region_from_named_shape(chart_slide, "CHART_AREA", default_chart_region)
    if placeholder2 is not None:
        _remove_shape(placeholder2)
    _insert_picture(chart_slide, chart_png, region2)

    out = Path(output_pptx)
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    return out
