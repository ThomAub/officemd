from __future__ import annotations

import json
from pathlib import Path

import pytest

from officemd import _cli
from officemd import (
    DocxPatch,
    DocxTextScope,
    MatchPolicy,
    PptxPatch,
    PptxTextScope,
    ScopedDocxReplace,
    ScopedPptxReplace,
    TextReplace,
    apply_ooxml_patch_json,
    detect_format,
    extract_csv_tables_ir_json,
    extract_ir_json,
    extract_sheet_names,
    extract_tables_ir_json,
    inspect_pdf_json,
    markdown_from_bytes,
    patch_docx,
    patch_pptx,
)

EXAMPLES_DIR = Path(__file__).resolve().parents[2]
DATA_DIR = EXAMPLES_DIR / "data"


def _read_bytes(filename: str) -> bytes:
    return (DATA_DIR / filename).read_bytes()


def test_showcase_format_detection() -> None:
    assert detect_format(_read_bytes("showcase.docx")) == ".docx"
    assert detect_format(_read_bytes("showcase.xlsx")) == ".xlsx"
    assert detect_format(_read_bytes("showcase.pptx")) == ".pptx"


def test_showcase_xlsx_sheet_names_and_tables_ir() -> None:
    xlsx = _read_bytes("showcase.xlsx")
    sheet_names = extract_sheet_names(xlsx)
    assert "Sales" in sheet_names
    assert "Summary" in sheet_names

    payload = extract_tables_ir_json(xlsx, style_aware_values=True, streaming_rows=True)
    doc = json.loads(payload)
    assert doc["kind"] == "Xlsx"
    assert len(doc["sheets"]) >= 2


def test_showcase_markdown_and_ir_smoke() -> None:
    docx = _read_bytes("showcase.docx")
    pptx = _read_bytes("showcase.pptx")

    docx_ir = json.loads(extract_ir_json(docx))
    pptx_ir = json.loads(extract_ir_json(pptx))
    assert docx_ir["kind"] == "Docx"
    assert pptx_ir["kind"] == "Pptx"

    markdown = markdown_from_bytes(pptx, format="pptx")
    assert "Quarterly Review" in markdown


def test_showcase_pptx_core_title_patch_surfaces_in_ir() -> None:
    patched = apply_ooxml_patch_json(
        _read_bytes("showcase.pptx"),
        json.dumps({"core_title": "Patched Showcase PPTX Title"}),
    )
    ir = json.loads(extract_ir_json(patched, format="pptx"))
    assert ir["kind"] == "Pptx"
    assert ir["properties"]["core"]["title"] == "Patched Showcase PPTX Title"


def test_typed_docx_patch_replaces_word_across_all_text() -> None:
    patched = patch_docx(
        _read_bytes("showcase.docx"),
        DocxPatch(
            set_core_title="Typed Showcase DOCX",
            scoped_replacements=[
                ScopedDocxReplace(
                    DocxTextScope.ALL_TEXT,
                    TextReplace("showcase", "demo", match_policy=MatchPolicy.CASE_INSENSITIVE),
                )
            ],
        ),
    )
    markdown = markdown_from_bytes(patched, format="docx", include_document_properties=True)
    assert "demo" in markdown.lower()
    ir = json.loads(extract_ir_json(patched, format="docx"))
    assert ir["properties"]["core"]["title"] == "Typed Showcase DOCX"


def test_typed_pptx_patch_replaces_word_across_all_text() -> None:
    patched = patch_pptx(
        _read_bytes("showcase.pptx"),
        PptxPatch(
            set_core_title="Typed Showcase PPTX",
            scoped_replacements=[
                ScopedPptxReplace(
                    PptxTextScope.ALL_TEXT,
                    TextReplace("Review", "Recap", match_policy=MatchPolicy.WHOLE_WORD),
                )
            ],
        ),
    )
    markdown = markdown_from_bytes(patched, format="pptx", include_document_properties=True)
    assert "Quarterly Recap" in markdown
    ir = json.loads(extract_ir_json(patched, format="pptx"))
    assert ir["properties"]["core"]["title"] == "Typed Showcase PPTX"


def test_showcase_csv_markdown_and_ir_smoke() -> None:
    csv_bytes = _read_bytes("showcase.csv")
    csv_ir = json.loads(extract_ir_json(csv_bytes, format="csv"))
    assert csv_ir["kind"] == "Xlsx"

    csv_tables_ir = json.loads(extract_csv_tables_ir_json(csv_bytes))
    assert csv_tables_ir["sheets"][0]["name"] == "Sheet1"

    markdown = markdown_from_bytes(csv_bytes, format="csv")
    assert "## Sheet: Sheet1" in markdown


def test_pdf_ocr_readiness_pair() -> None:
    scanned = json.loads(inspect_pdf_json(_read_bytes("ocr_graph_scanned.pdf")))
    ocred = json.loads(inspect_pdf_json(_read_bytes("ocr_graph_ocred.pdf")))

    assert scanned["classification"] in {"Scanned", "ImageBased", "Mixed"}
    assert scanned["pages_needing_ocr"]
    assert ocred["classification"] == "TextBased"
    assert not ocred["pages_needing_ocr"]


def test_pdf_textbased_fixture_has_markdown_without_ocr() -> None:
    textbased = _read_bytes("ocr_tagged_textbased.pdf")
    diagnostics = json.loads(inspect_pdf_json(textbased))

    assert diagnostics["classification"] == "TextBased"
    assert not diagnostics["pages_needing_ocr"]

    markdown = markdown_from_bytes(textbased, format="pdf")
    assert "## Page: 1" in markdown
    assert "Heading" in markdown


def test_cli_pdf_page_selection_single_and_range() -> None:
    pdf = _read_bytes("ocr_tagged_textbased.pdf")
    markdown = markdown_from_bytes(pdf, format="pdf")

    page_two = _cli._apply_selectors(
        markdown,
        doc_format="pdf",
        pages_selector="2",
        sheets_selector=None,
    )
    assert "## Page: 2" in page_two
    assert "## Page: 1" not in page_two

    page_range = _cli._apply_selectors(
        markdown,
        doc_format="pdf",
        pages_selector="1-2",
        sheets_selector=None,
    )
    assert "## Page: 1" in page_range
    assert "## Page: 2" in page_range


def test_cli_xlsx_sheet_selection_by_name_index_and_range() -> None:
    xlsx = _read_bytes("showcase.xlsx")
    markdown = markdown_from_bytes(xlsx, format="xlsx")

    by_name = _cli._apply_selectors(
        markdown,
        doc_format="xlsx",
        pages_selector=None,
        sheets_selector="Summary",
    )
    assert "## Sheet: Summary" in by_name
    assert "## Sheet: Sales" not in by_name

    by_index = _cli._apply_selectors(
        markdown,
        doc_format="xlsx",
        pages_selector=None,
        sheets_selector="1",
    )
    assert "## Sheet: Sales" in by_index
    assert "## Sheet: Summary" not in by_index

    by_range = _cli._apply_selectors(
        markdown,
        doc_format="xlsx",
        pages_selector=None,
        sheets_selector="1-2",
    )
    assert "## Sheet: Sales" in by_range
    assert "## Sheet: Summary" in by_range

    by_pages = _cli._apply_selectors(
        markdown,
        doc_format="xlsx",
        pages_selector="2",
        sheets_selector=None,
    )
    assert "## Sheet: Summary" in by_pages
    assert "## Sheet: Sales" not in by_pages

    combined = _cli._apply_selectors(
        markdown,
        doc_format="xlsx",
        pages_selector="2",
        sheets_selector="Sales",
    )
    assert "## Sheet: Summary" in combined
    assert "## Sheet: Sales" in combined


def test_cli_selector_invalid_handling() -> None:
    with pytest.raises(_cli.CliUsageError, match="start must be <= end"):
        _cli._parse_numeric_selector("4-2", flag="--pages")

    with pytest.raises(_cli.CliUsageError, match="supported for PDF/PPTX and XLSX/CSV"):
        _cli._apply_selectors(
            "# Title\n",
            doc_format="docx",
            pages_selector="1",
            sheets_selector=None,
        )

    with pytest.raises(_cli.CliUsageError, match="only supported for XLSX/CSV"):
        _cli._apply_selectors(
            "## Page: 1\ncontent\n",
            doc_format="pdf",
            pages_selector=None,
            sheets_selector="1",
        )

    with pytest.raises(_cli.CliUsageError, match="did not match any sheet headings"):
        _cli._apply_selectors(
            "## Sheet: Sales\ncontent\n",
            doc_format="xlsx",
            pages_selector=None,
            sheets_selector="Missing",
        )
