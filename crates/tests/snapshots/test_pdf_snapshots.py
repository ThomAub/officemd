from __future__ import annotations

import json
from pathlib import Path

from officemd import extract_ir_json, inspect_pdf_json, markdown_from_bytes

ROOT = Path(__file__).resolve().parents[3]
DATA_DIR = ROOT / "examples" / "data"

PDF_FIXTURES = {
    "openxml_whitepaper": "OpenXML_WhitePaper.pdf",
    "ocr_graph_ocred": "ocr_graph_ocred.pdf",
    "ocr_graph_scanned": "ocr_graph_scanned.pdf",
    "ocr_tagged_textbased": "ocr_tagged_textbased.pdf",
    "encoding_heuristic": "encoding_heuristic_fixture.pdf",
}


def _read_fixture(name: str) -> bytes:
    path = DATA_DIR / name
    if not path.exists():
        raise AssertionError(f"missing fixture: {path}")
    return path.read_bytes()


def _canonical_json(payload: str) -> str:
    return json.dumps(json.loads(payload), indent=2, sort_keys=True, ensure_ascii=False) + "\n"


def _canonical_markdown(payload: str | None) -> str:
    if not payload:
        return ""
    return payload.replace("\r\n", "\n").rstrip() + "\n"


# ---------------------------------------------------------------------------
# OpenXML_WhitePaper.pdf
# ---------------------------------------------------------------------------


def test_openxml_whitepaper_ir_snapshot(file_regression) -> None:
    content = _read_fixture(PDF_FIXTURES["openxml_whitepaper"])
    payload = extract_ir_json(content, format="pdf")
    file_regression.check(
        _canonical_json(payload), extension=".json", basename="openxml_whitepaper_ir"
    )


def test_openxml_whitepaper_markdown_snapshot(file_regression) -> None:
    content = _read_fixture(PDF_FIXTURES["openxml_whitepaper"])
    payload = markdown_from_bytes(content, format="pdf")
    file_regression.check(
        _canonical_markdown(payload),
        extension=".md",
        basename="openxml_whitepaper_markdown",
    )


def test_openxml_whitepaper_inspect_snapshot(file_regression) -> None:
    content = _read_fixture(PDF_FIXTURES["openxml_whitepaper"])
    payload = inspect_pdf_json(content)
    file_regression.check(
        _canonical_json(payload), extension=".json", basename="openxml_whitepaper_inspect"
    )


# ---------------------------------------------------------------------------
# ocr_graph_ocred.pdf
# ---------------------------------------------------------------------------


def test_ocr_graph_ocred_ir_snapshot(file_regression) -> None:
    content = _read_fixture(PDF_FIXTURES["ocr_graph_ocred"])
    payload = extract_ir_json(content, format="pdf")
    file_regression.check(
        _canonical_json(payload), extension=".json", basename="ocr_graph_ocred_ir"
    )


def test_ocr_graph_ocred_markdown_snapshot(file_regression) -> None:
    content = _read_fixture(PDF_FIXTURES["ocr_graph_ocred"])
    payload = markdown_from_bytes(content, format="pdf")
    file_regression.check(
        _canonical_markdown(payload),
        extension=".md",
        basename="ocr_graph_ocred_markdown",
    )


def test_ocr_graph_ocred_inspect_snapshot(file_regression) -> None:
    content = _read_fixture(PDF_FIXTURES["ocr_graph_ocred"])
    payload = inspect_pdf_json(content)
    file_regression.check(
        _canonical_json(payload), extension=".json", basename="ocr_graph_ocred_inspect"
    )


# ---------------------------------------------------------------------------
# ocr_graph_scanned.pdf
# ---------------------------------------------------------------------------


def test_ocr_graph_scanned_ir_snapshot(file_regression) -> None:
    content = _read_fixture(PDF_FIXTURES["ocr_graph_scanned"])
    payload = extract_ir_json(content, format="pdf")
    file_regression.check(
        _canonical_json(payload), extension=".json", basename="ocr_graph_scanned_ir"
    )


def test_ocr_graph_scanned_markdown_snapshot(file_regression) -> None:
    content = _read_fixture(PDF_FIXTURES["ocr_graph_scanned"])
    payload = markdown_from_bytes(content, format="pdf")
    file_regression.check(
        _canonical_markdown(payload),
        extension=".md",
        basename="ocr_graph_scanned_markdown",
    )


def test_ocr_graph_scanned_inspect_snapshot(file_regression) -> None:
    content = _read_fixture(PDF_FIXTURES["ocr_graph_scanned"])
    payload = inspect_pdf_json(content)
    file_regression.check(
        _canonical_json(payload), extension=".json", basename="ocr_graph_scanned_inspect"
    )


# ---------------------------------------------------------------------------
# ocr_tagged_textbased.pdf
# ---------------------------------------------------------------------------


def test_ocr_tagged_textbased_ir_snapshot(file_regression) -> None:
    content = _read_fixture(PDF_FIXTURES["ocr_tagged_textbased"])
    payload = extract_ir_json(content, format="pdf")
    file_regression.check(
        _canonical_json(payload),
        extension=".json",
        basename="ocr_tagged_textbased_ir",
    )


def test_ocr_tagged_textbased_markdown_snapshot(file_regression) -> None:
    content = _read_fixture(PDF_FIXTURES["ocr_tagged_textbased"])
    payload = markdown_from_bytes(content, format="pdf")
    file_regression.check(
        _canonical_markdown(payload),
        extension=".md",
        basename="ocr_tagged_textbased_markdown",
    )


def test_ocr_tagged_textbased_inspect_snapshot(file_regression) -> None:
    content = _read_fixture(PDF_FIXTURES["ocr_tagged_textbased"])
    payload = inspect_pdf_json(content)
    file_regression.check(
        _canonical_json(payload),
        extension=".json",
        basename="ocr_tagged_textbased_inspect",
    )


# ---------------------------------------------------------------------------
# encoding_heuristic_fixture.pdf
# ---------------------------------------------------------------------------


def test_encoding_heuristic_ir_snapshot(file_regression) -> None:
    content = _read_fixture(PDF_FIXTURES["encoding_heuristic"])
    payload = extract_ir_json(content, format="pdf")
    file_regression.check(
        _canonical_json(payload),
        extension=".json",
        basename="encoding_heuristic_ir",
    )


def test_encoding_heuristic_markdown_snapshot(file_regression) -> None:
    content = _read_fixture(PDF_FIXTURES["encoding_heuristic"])
    payload = markdown_from_bytes(content, format="pdf")
    file_regression.check(
        _canonical_markdown(payload),
        extension=".md",
        basename="encoding_heuristic_markdown",
    )


def test_encoding_heuristic_inspect_snapshot(file_regression) -> None:
    content = _read_fixture(PDF_FIXTURES["encoding_heuristic"])
    payload = inspect_pdf_json(content)
    file_regression.check(
        _canonical_json(payload),
        extension=".json",
        basename="encoding_heuristic_inspect",
    )
