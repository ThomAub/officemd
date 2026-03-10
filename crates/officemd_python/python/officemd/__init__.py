from officemd._officemd import (
    detect_format,
    docling_from_bytes,
    extract_csv_tables_ir_json,
    extract_ir_json,
    extract_sheet_names,
    extract_tables_ir_json,
    inspect_pdf_fonts_json,
    inspect_pdf_json,
    markdown_from_bytes_batch,
    markdown_from_bytes,
)

def _missing_rich(*_args, **_kwargs):
    raise ModuleNotFoundError(
        "rich is required for render/diff helpers. Install with: pip install rich"
    )


try:
    from officemd.diff import diff_markdown, render_diff
    from officemd.render import render_markdown
except ModuleNotFoundError:
    diff_markdown = _missing_rich
    render_diff = _missing_rich
    render_markdown = _missing_rich

__all__ = [
    "detect_format",
    "diff_markdown",
    "docling_from_bytes",
    "extract_csv_tables_ir_json",
    "extract_ir_json",
    "extract_sheet_names",
    "extract_tables_ir_json",
    "inspect_pdf_fonts_json",
    "inspect_pdf_json",
    "markdown_from_bytes_batch",
    "markdown_from_bytes",
    "render_diff",
    "render_markdown",
]

__version__ = "0.1.1"
