"""Compare officemd markdown output against markitdown and docling.

These tests require the `bench` dependency group:
    uv sync --group bench
    uv run maturin develop --release
    uv run pytest ../../crates/tests/benchmarks/ -v

Tests skip gracefully when markitdown or docling are not installed.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from officemd import markdown_from_bytes

ROOT = Path(__file__).resolve().parents[3]
DATA_DIR = ROOT / "examples" / "data"

DOCX_FILE = DATA_DIR / "showcase.docx"
XLSX_FILE = DATA_DIR / "showcase.xlsx"
PPTX_FILE = DATA_DIR / "showcase.pptx"
PDF_FILE = DATA_DIR / "OpenXML_WhitePaper.pdf"

FIXTURES = {
    "docx": DOCX_FILE,
    "xlsx": XLSX_FILE,
    "pptx": PPTX_FILE,
    "pdf": PDF_FILE,
}


def _read(path: Path) -> bytes:
    if not path.exists():
        pytest.skip(f"fixture missing: {path.name}")
    return path.read_bytes()


# ---------------------------------------------------------------------------
# markitdown comparison
# ---------------------------------------------------------------------------

markitdown_mod = pytest.importorskip("markitdown", reason="markitdown not installed (uv sync --group bench)")


@pytest.fixture(scope="module")
def markitdown_converter():
    from markitdown import MarkItDown
    return MarkItDown()


@pytest.mark.parametrize("fmt", ["docx", "xlsx", "pptx"])
def test_markitdown_produces_output(fmt, markitdown_converter):
    """Both officemd and markitdown produce non-empty markdown."""
    path = FIXTURES[fmt]
    content = _read(path)

    officemd_md = markdown_from_bytes(content, format=fmt)
    markitdown_md = markitdown_converter.convert(str(path)).text_content

    assert len(officemd_md.strip()) > 0, "officemd returned empty markdown"
    assert len(markitdown_md.strip()) > 0, "markitdown returned empty markdown"


@pytest.mark.parametrize("fmt", ["docx", "xlsx", "pptx"])
def test_officemd_shorter_than_markitdown(fmt, markitdown_converter):
    """officemd compact markdown should be shorter or equal in length."""
    path = FIXTURES[fmt]
    content = _read(path)

    officemd_md = markdown_from_bytes(content, format=fmt)
    markitdown_md = markitdown_converter.convert(str(path)).text_content

    officemd_len = len(officemd_md)
    markitdown_len = len(markitdown_md)

    print(f"\n{fmt}: officemd={officemd_len} chars, markitdown={markitdown_len} chars")
    # Not a hard assertion - just reporting. officemd aims for compact output
    # but markitdown may omit content that officemd includes.


# ---------------------------------------------------------------------------
# docling comparison
# ---------------------------------------------------------------------------

docling_mod = pytest.importorskip("docling", reason="docling not installed (uv sync --group bench)")


@pytest.fixture(scope="module")
def docling_converter():
    from docling.document_converter import DocumentConverter
    return DocumentConverter()


@pytest.mark.parametrize("fmt", ["docx", "xlsx", "pptx"])
def test_docling_produces_output(fmt, docling_converter):
    """Both officemd and docling produce non-empty markdown."""
    path = FIXTURES[fmt]
    content = _read(path)

    officemd_md = markdown_from_bytes(content, format=fmt)
    docling_md = docling_converter.convert(str(path)).document.export_to_markdown()

    assert len(officemd_md.strip()) > 0, "officemd returned empty markdown"
    assert len(docling_md.strip()) > 0, "docling returned empty markdown"


@pytest.mark.parametrize("fmt", ["docx", "xlsx", "pptx"])
def test_output_char_comparison(fmt, markitdown_converter, docling_converter):
    """Side-by-side character count comparison across all three tools."""
    path = FIXTURES[fmt]
    content = _read(path)

    officemd_md = markdown_from_bytes(content, format=fmt)
    markitdown_md = markitdown_converter.convert(str(path)).text_content
    docling_md = docling_converter.convert(str(path)).document.export_to_markdown()

    print(f"\n{fmt} character counts:")
    print(f"  officemd:   {len(officemd_md):>8}")
    print(f"  markitdown: {len(markitdown_md):>8}")
    print(f"  docling:    {len(docling_md):>8}")
