from __future__ import annotations

from pathlib import Path

import pandas as pd

from .io_excel import read_excel_to_df
from .utils import ensure_dir, safe_filename


def _next_available_path(path: Path) -> Path:
    """Return a non-conflicting path by appending _N when needed."""
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


def _normalize_ctx(df: pd.DataFrame, row: pd.Series) -> dict[str, object]:
    """Convert a dataframe row into a template context dict."""
    ctx: dict[str, object] = {}
    for col in df.columns:
        val = row[col]
        if pd.isna(val):
            ctx[col] = ""
        else:
            ctx[col] = val

    # Common derived fields
    if "Total" not in ctx and "Qty" in ctx and "UnitPrice" in ctx:
        try:
            ctx["Total"] = float(ctx["Qty"]) * float(ctx["UnitPrice"])
        except Exception:
            pass

    return ctx


def _replace_in_paragraph_runs(paragraph, mapping: dict[str, str]) -> None:
    # Best-effort: replace within each run (keeps formatting).
    # NOTE: If Word splits a placeholder across multiple runs, this won't catch it.
    for run in paragraph.runs:
        text = run.text
        if not text:
            continue
        for k, v in mapping.items():
            text = text.replace(f"{{{{ {k} }}}}", v).replace(f"{{{{{k}}}}}", v)
        run.text = text


def _replace_in_doc(doc, mapping: dict[str, str]) -> None:
    # Body paragraphs
    for p in doc.paragraphs:
        _replace_in_paragraph_runs(p, mapping)

    # Tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _replace_in_paragraph_runs(p, mapping)

    # Headers/footers (common)
    for section in doc.sections:
        for p in section.header.paragraphs:
            _replace_in_paragraph_runs(p, mapping)
        for table in section.header.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        _replace_in_paragraph_runs(p, mapping)

        for p in section.footer.paragraphs:
            _replace_in_paragraph_runs(p, mapping)
        for table in section.footer.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        _replace_in_paragraph_runs(p, mapping)


def _render_with_template(template_docx: Path, ctx: dict[str, object], out_path: Path) -> None:
    """Render a DOCX template.

    Priority:
      1) If `docxtpl` is installed, use it (more robust templating).
      2) Otherwise, do a simple {{var}} text replacement using python-docx.
    """
    # 1) Try docxtpl when available.
    # If docxtpl is installed but rendering fails, surface the error.
    try:
        from docxtpl import DocxTemplate  # type: ignore
    except ImportError:
        DocxTemplate = None  # type: ignore[assignment]

    if DocxTemplate is not None:
        tpl = DocxTemplate(str(template_docx))
        tpl.render(ctx)
        tpl.save(str(out_path))
        return

    # 2) Simple replacement (no extra deps)
    from docx import Document  # lazy import

    doc = Document(str(template_docx))
    mapping = {k: "" if v is None else str(v) for k, v in ctx.items()}
    _replace_in_doc(doc, mapping)
    doc.save(str(out_path))


def _render_simple_docx(ctx: dict[str, object], out_path: Path, title_template: str) -> None:
    """Fallback renderer using python-docx (no template)."""
    from docx import Document  # imported lazily

    doc = Document()
    try:
        title = title_template.format(**{k: str(v) for k, v in ctx.items()})
    except Exception:
        title = "Row Document"
    doc.add_heading(title, level=1)

    doc.add_paragraph("This document was generated automatically from an Excel row.")

    keys = list(ctx.keys())
    preferred_order = [k for k in ("ID", "Name", "Department", "Category", "OrderDate", "Email", "Qty", "UnitPrice", "Total") if k in ctx]
    rest = [k for k in keys if k not in preferred_order]
    ordered = preferred_order + rest

    table = doc.add_table(rows=len(ordered), cols=2)
    table.style = "Table Grid"
    for i, k in enumerate(ordered):
        table.cell(i, 0).text = str(k)
        table.cell(i, 1).text = str(ctx.get(k, ""))

    doc.save(str(out_path))


def create_word_docs_from_excel(
    input_xlsx: str | Path,
    output_dir: str | Path,
    sheet_name: int | str = 0,
    template_docx: str | Path | None = None,
    filename_cols: tuple[str, ...] = ("ID", "Name"),
    title_template: str = "Order Summary - {ID} / {Name}",
) -> list[Path]:
    """Create one DOCX per row from an Excel sheet.

    - If `template_docx` is provided: render the template.
    - Otherwise: generate a simple doc using python-docx.

    Template variables: any Excel column name becomes a variable, e.g. {{ID}}, {{Name}} ...
    Derived variables: if Qty + UnitPrice exist, Total is auto-computed when missing.
    """
    df = read_excel_to_df(input_xlsx, sheet_name=sheet_name)
    out_dir = ensure_dir(output_dir)
    out_paths: list[Path] = []

    tpl_path: Path | None = Path(template_docx) if template_docx else None
    if tpl_path is not None and not tpl_path.exists():
        raise FileNotFoundError(f"Template not found: {tpl_path}")

    for _, row in df.iterrows():
        ctx = _normalize_ctx(df, row)

        parts: list[str] = []
        for col in filename_cols:
            if col in ctx and str(ctx[col]).strip():
                parts.append(str(ctx[col]))
        fname = safe_filename("_".join(parts) if parts else "row") + ".docx"
        out_path = _next_available_path(out_dir / fname)

        if tpl_path is not None:
            _render_with_template(tpl_path, ctx, out_path)
        else:
            _render_simple_docx(ctx, out_path, title_template=title_template)

        out_paths.append(out_path)

    return out_paths
