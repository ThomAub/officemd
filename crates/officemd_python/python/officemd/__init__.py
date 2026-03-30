from officemd._officemd import (  # type: ignore[unresolved-import]
    apply_ooxml_patch_json,
    detect_format,
    docling_from_bytes,
    extract_csv_tables_ir_json,
    extract_ir_json,
    extract_sheet_names,
    extract_tables_ir_json,
    inspect_pdf_fonts_json,
    inspect_pdf_json,
    markdown_from_bytes,
    markdown_from_bytes_batch,
)
from officemd.patching import (
    BatchPatchContentResult,
    BatchPatchJob,
    BatchPatchResult,
    PatchContentResult,
    DocxPatch,
    DocxTextScope,
    MatchPolicy,
    PptxPatch,
    PptxTextScope,
    ReplaceMode,
    ScopedDocxReplace,
    ScopedPptxReplace,
    ScopedXlsxReplace,
    TextReplace,
    PatchReport,
    XlsxPatch,
    XlsxSheetRename,
    XlsxTextScope,
    patch_docx,
    patch_docx_batch,
    patch_docx_batch_with_report,
    patch_docx_with_report,
    patch_files,
    patch_pptx,
    patch_pptx_batch,
    patch_pptx_batch_with_report,
    patch_pptx_with_report,
    patch_xlsx,
    patch_xlsx_batch,
    patch_xlsx_batch_with_report,
    patch_xlsx_with_report,
)


def _missing_rich(*_args, **_kwargs):
    raise ModuleNotFoundError(
        "rich is required for render/diff helpers. Install with: pip install rich"
    )


try:
    from officemd.diff import diff_markdown, render_diff
    from officemd.render import render_markdown
except ModuleNotFoundError:
    diff_markdown = _missing_rich  # type: ignore[invalid-assignment]
    render_diff = _missing_rich  # type: ignore[invalid-assignment]
    render_markdown = _missing_rich  # type: ignore[invalid-assignment]

__all__ = [
    "apply_ooxml_patch_json",
    "BatchPatchContentResult",
    "BatchPatchJob",
    "BatchPatchResult",
    "PatchContentResult",
    "detect_format",
    "diff_markdown",
    "DocxPatch",
    "DocxTextScope",
    "docling_from_bytes",
    "extract_csv_tables_ir_json",
    "extract_ir_json",
    "extract_sheet_names",
    "extract_tables_ir_json",
    "inspect_pdf_fonts_json",
    "MatchPolicy",
    "inspect_pdf_json",
    "markdown_from_bytes",
    "PatchReport",
    "patch_docx",
    "patch_docx_batch",
    "patch_docx_batch_with_report",
    "patch_docx_with_report",
    "patch_files",
    "patch_pptx",
    "patch_pptx_batch",
    "patch_pptx_batch_with_report",
    "patch_pptx_with_report",
    "PptxPatch",
    "PptxTextScope",
    "markdown_from_bytes_batch",
    "ReplaceMode",
    "render_diff",
    "render_markdown",
    "ScopedDocxReplace",
    "ScopedPptxReplace",
    "ScopedXlsxReplace",
    "TextReplace",
    "XlsxPatch",
    "XlsxSheetRename",
    "XlsxTextScope",
    "patch_xlsx",
    "patch_xlsx_batch",
    "patch_xlsx_batch_with_report",
    "patch_xlsx_with_report",
]

__version__ = "0.1.1"
