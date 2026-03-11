# /// script
# requires-python = ">=3.12"
# dependencies = [
#     "officemd",
#     "markitdown[all]==0.1.5",
#     "docling==2.78.0",
# ]
# ///
"""Generate cross-tool comparison markdown snapshots.

Produces markdown output from officemd, markitdown, and docling for all
fixtures, written to crates/tests/snapshots/cross_tool/<fixture_name>/.

Run manually:
    uv run crates/tests/snapshots/generate_cross_tool_snapshots.py

These files are committed for git diff review but never asserted in CI.
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
DATA_DIR = ROOT / "examples" / "data"
OUTPUT_DIR = Path(__file__).resolve().parent / "cross_tool"

FIXTURES = {
    "showcase_docx": ("showcase.docx", "docx"),
    "showcase_02_docx": ("showcase_02.docx", "docx"),
    "showcase_xlsx": ("showcase.xlsx", "xlsx"),
    "trim_sparse_trailing_xlsx": ("trim_sparse_trailing.xlsx", "xlsx"),
    "trim_wide_sparse_xlsx": ("trim_wide_sparse.xlsx", "xlsx"),
    "trim_all_empty_xlsx": ("trim_all_empty.xlsx", "xlsx"),
    "showcase_csv": ("showcase.csv", "csv"),
    "showcase_pptx": ("showcase.pptx", "pptx"),
    "openxml_whitepaper_pdf": ("OpenXML_WhitePaper.pdf", "pdf"),
    "ocr_graph_ocred_pdf": ("ocr_graph_ocred.pdf", "pdf"),
    "ocr_graph_scanned_pdf": ("ocr_graph_scanned.pdf", "pdf"),
    "ocr_tagged_textbased_pdf": ("ocr_tagged_textbased.pdf", "pdf"),
    "encoding_heuristic_pdf": ("encoding_heuristic_fixture.pdf", "pdf"),
}

# markitdown format support
MARKITDOWN_FORMATS = {"docx", "xlsx", "pptx", "pdf", "csv"}
# docling format support
DOCLING_FORMATS = {"docx", "xlsx", "pptx", "pdf"}


def generate_officemd(content: bytes, fmt: str) -> str:
    from officemd import markdown_from_bytes

    try:
        return markdown_from_bytes(content, format=fmt)
    except Exception as exc:
        return f"<!-- ERROR: {exc} -->"


def generate_markitdown(filepath: Path) -> str:
    try:
        from markitdown import MarkItDown

        converter = MarkItDown()
        result = converter.convert(str(filepath))
        return result.text_content
    except Exception as exc:
        return f"<!-- ERROR: {exc} -->"


def generate_docling(filepath: Path) -> str:
    try:
        from docling.document_converter import DocumentConverter

        converter = DocumentConverter()
        result = converter.convert(str(filepath))
        return result.document.export_to_markdown()
    except Exception as exc:
        return f"<!-- ERROR: {exc} -->"


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    for fixture_name, (filename, fmt) in FIXTURES.items():
        filepath = DATA_DIR / filename
        if not filepath.exists():
            print(f"SKIP {fixture_name}: {filepath} not found", file=sys.stderr)
            continue

        outdir = OUTPUT_DIR / fixture_name
        outdir.mkdir(parents=True, exist_ok=True)

        content = filepath.read_bytes()
        print(f"Processing {fixture_name}...")

        # officemd
        md = generate_officemd(content, fmt)
        (outdir / "officemd.md").write_text(md, encoding="utf-8")

        # markitdown
        if fmt in MARKITDOWN_FORMATS:
            md = generate_markitdown(filepath)
        else:
            md = "<!-- NOT SUPPORTED -->"
        (outdir / "markitdown.md").write_text(md, encoding="utf-8")

        # docling
        if fmt in DOCLING_FORMATS:
            md = generate_docling(filepath)
        else:
            md = "<!-- NOT SUPPORTED -->"
        (outdir / "docling.md").write_text(md, encoding="utf-8")

    print("Done. Output in:", OUTPUT_DIR)


if __name__ == "__main__":
    main()
