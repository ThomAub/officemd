from __future__ import annotations

import json
from pathlib import Path

from officemd import docling_from_bytes

ROOT = Path(__file__).resolve().parents[3]
DATA_DIR = ROOT / "examples" / "data"


def _read_fixture(name: str) -> bytes:
    path = DATA_DIR / name
    if not path.exists():
        raise AssertionError(f"missing fixture: {path}")
    return path.read_bytes()


def _canonical_json(payload: str) -> str:
    return json.dumps(json.loads(payload), indent=2, sort_keys=True, ensure_ascii=False) + "\n"


def test_showcase_docx_docling_snapshot(file_regression) -> None:
    content = _read_fixture("showcase.docx")
    payload = docling_from_bytes(content, format="docx")
    file_regression.check(
        _canonical_json(payload), extension=".json", basename="showcase_docx_docling"
    )


def test_showcase_xlsx_docling_snapshot(file_regression) -> None:
    content = _read_fixture("showcase.xlsx")
    payload = docling_from_bytes(content, format="xlsx")
    file_regression.check(
        _canonical_json(payload), extension=".json", basename="showcase_xlsx_docling"
    )


def test_showcase_pptx_docling_snapshot(file_regression) -> None:
    content = _read_fixture("showcase.pptx")
    payload = docling_from_bytes(content, format="pptx")
    file_regression.check(
        _canonical_json(payload), extension=".json", basename="showcase_pptx_docling"
    )


def test_showcase_csv_docling_snapshot(file_regression) -> None:
    content = _read_fixture("showcase.csv")
    payload = docling_from_bytes(content, format="csv")
    file_regression.check(
        _canonical_json(payload), extension=".json", basename="showcase_csv_docling"
    )


def test_showcase_02_docx_docling_snapshot(file_regression) -> None:
    content = _read_fixture("showcase_02.docx")
    payload = docling_from_bytes(content, format="docx")
    file_regression.check(
        _canonical_json(payload), extension=".json", basename="showcase_02_docx_docling"
    )


def test_openxml_whitepaper_pdf_docling_snapshot(file_regression) -> None:
    content = _read_fixture("OpenXML_WhitePaper.pdf")
    payload = docling_from_bytes(content, format="pdf")
    file_regression.check(
        _canonical_json(payload),
        extension=".json",
        basename="openxml_whitepaper_pdf_docling",
    )


def test_ocr_tagged_textbased_pdf_docling_snapshot(file_regression) -> None:
    content = _read_fixture("ocr_tagged_textbased.pdf")
    payload = docling_from_bytes(content, format="pdf")
    file_regression.check(
        _canonical_json(payload),
        extension=".json",
        basename="ocr_tagged_textbased_pdf_docling",
    )


def test_ocr_graph_ocred_pdf_docling_snapshot(file_regression) -> None:
    content = _read_fixture("ocr_graph_ocred.pdf")
    payload = docling_from_bytes(content, format="pdf")
    file_regression.check(
        _canonical_json(payload),
        extension=".json",
        basename="ocr_graph_ocred_pdf_docling",
    )
