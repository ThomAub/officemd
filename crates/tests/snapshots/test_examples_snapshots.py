from __future__ import annotations

import json
from pathlib import Path

from officemd import extract_ir_json, extract_tables_ir_json, markdown_from_bytes

ROOT = Path(__file__).resolve().parents[3]
DATA_DIR = ROOT / "examples" / "data"


def _read_fixture(name: str) -> bytes:
    path = DATA_DIR / name
    if not path.exists():
        raise AssertionError(f"missing fixture: {path}")
    return path.read_bytes()


def _canonical_json(payload: str) -> str:
    return json.dumps(json.loads(payload), indent=2, sort_keys=True, ensure_ascii=False) + "\n"


def _canonical_markdown(payload: str) -> str:
    return payload.replace("\r\n", "\n").rstrip() + "\n"


def test_showcase_docx_ir_snapshot(file_regression) -> None:
    content = _read_fixture("showcase.docx")
    payload = extract_ir_json(content, format="docx")
    file_regression.check(_canonical_json(payload), extension=".json", basename="showcase_docx_ir")


def test_showcase_docx_markdown_snapshot(file_regression) -> None:
    content = _read_fixture("showcase.docx")
    payload = markdown_from_bytes(content, format="docx")
    file_regression.check(
        _canonical_markdown(payload),
        extension=".md",
        basename="showcase_docx_markdown",
    )


def test_showcase_docx_markdown_with_properties_snapshot(file_regression) -> None:
    content = _read_fixture("showcase.docx")
    payload = markdown_from_bytes(
        content,
        format="docx",
        include_document_properties=True,
    )
    file_regression.check(
        _canonical_markdown(payload),
        extension=".md",
        basename="showcase_docx_markdown_with_properties",
    )


def test_showcase_xlsx_ir_snapshot(file_regression) -> None:
    content = _read_fixture("showcase.xlsx")
    payload = extract_ir_json(content, format="xlsx")
    file_regression.check(_canonical_json(payload), extension=".json", basename="showcase_xlsx_ir")


def test_showcase_xlsx_markdown_snapshot(file_regression) -> None:
    content = _read_fixture("showcase.xlsx")
    payload = markdown_from_bytes(content, format="xlsx")
    file_regression.check(
        _canonical_markdown(payload),
        extension=".md",
        basename="showcase_xlsx_markdown",
    )


def test_showcase_xlsx_markdown_with_properties_snapshot(file_regression) -> None:
    content = _read_fixture("showcase.xlsx")
    payload = markdown_from_bytes(
        content,
        format="xlsx",
        include_document_properties=True,
    )
    file_regression.check(
        _canonical_markdown(payload),
        extension=".md",
        basename="showcase_xlsx_markdown_with_properties",
    )


def test_showcase_xlsx_tables_streaming_style_aware_snapshot(file_regression) -> None:
    content = _read_fixture("showcase.xlsx")
    payload = extract_tables_ir_json(content, style_aware_values=True, streaming_rows=True)
    file_regression.check(
        _canonical_json(payload),
        extension=".json",
        basename="showcase_xlsx_tables_streaming_style_aware",
    )


def test_showcase_pptx_ir_snapshot(file_regression) -> None:
    content = _read_fixture("showcase.pptx")
    payload = extract_ir_json(content, format="pptx")
    file_regression.check(_canonical_json(payload), extension=".json", basename="showcase_pptx_ir")


def test_showcase_pptx_markdown_snapshot(file_regression) -> None:
    content = _read_fixture("showcase.pptx")
    payload = markdown_from_bytes(content, format="pptx")
    file_regression.check(
        _canonical_markdown(payload),
        extension=".md",
        basename="showcase_pptx_markdown",
    )


def test_showcase_pptx_markdown_with_properties_snapshot(file_regression) -> None:
    content = _read_fixture("showcase.pptx")
    payload = markdown_from_bytes(
        content,
        format="pptx",
        include_document_properties=True,
    )
    file_regression.check(
        _canonical_markdown(payload),
        extension=".md",
        basename="showcase_pptx_markdown_with_properties",
    )


# ---------------------------------------------------------------------------
# showcase_02.docx
# ---------------------------------------------------------------------------


def test_showcase_02_docx_ir_snapshot(file_regression) -> None:
    content = _read_fixture("showcase_02.docx")
    payload = extract_ir_json(content, format="docx")
    file_regression.check(
        _canonical_json(payload), extension=".json", basename="showcase_02_docx_ir"
    )


def test_showcase_02_docx_markdown_snapshot(file_regression) -> None:
    content = _read_fixture("showcase_02.docx")
    payload = markdown_from_bytes(content, format="docx")
    file_regression.check(
        _canonical_markdown(payload),
        extension=".md",
        basename="showcase_02_docx_markdown",
    )


# ---------------------------------------------------------------------------
# showcase.csv
# ---------------------------------------------------------------------------


def test_showcase_csv_ir_snapshot(file_regression) -> None:
    content = _read_fixture("showcase.csv")
    payload = extract_ir_json(content, format="csv")
    file_regression.check(
        _canonical_json(payload), extension=".json", basename="showcase_csv_ir"
    )


def test_showcase_csv_markdown_snapshot(file_regression) -> None:
    content = _read_fixture("showcase.csv")
    payload = markdown_from_bytes(content, format="csv")
    file_regression.check(
        _canonical_markdown(payload),
        extension=".md",
        basename="showcase_csv_markdown",
    )


# ---------------------------------------------------------------------------
# edge-case XLSX fixtures
# ---------------------------------------------------------------------------


def test_trim_sparse_trailing_xlsx_ir_snapshot(file_regression) -> None:
    content = _read_fixture("trim_sparse_trailing.xlsx")
    payload = extract_ir_json(content, format="xlsx")
    file_regression.check(
        _canonical_json(payload),
        extension=".json",
        basename="trim_sparse_trailing_xlsx_ir",
    )


def test_trim_sparse_trailing_xlsx_markdown_snapshot(file_regression) -> None:
    content = _read_fixture("trim_sparse_trailing.xlsx")
    payload = markdown_from_bytes(content, format="xlsx")
    file_regression.check(
        _canonical_markdown(payload),
        extension=".md",
        basename="trim_sparse_trailing_xlsx_markdown",
    )


def test_trim_wide_sparse_xlsx_ir_snapshot(file_regression) -> None:
    content = _read_fixture("trim_wide_sparse.xlsx")
    payload = extract_ir_json(content, format="xlsx")
    file_regression.check(
        _canonical_json(payload),
        extension=".json",
        basename="trim_wide_sparse_xlsx_ir",
    )


def test_trim_wide_sparse_xlsx_markdown_snapshot(file_regression) -> None:
    content = _read_fixture("trim_wide_sparse.xlsx")
    payload = markdown_from_bytes(content, format="xlsx")
    file_regression.check(
        _canonical_markdown(payload),
        extension=".md",
        basename="trim_wide_sparse_xlsx_markdown",
    )


def test_trim_all_empty_xlsx_ir_snapshot(file_regression) -> None:
    content = _read_fixture("trim_all_empty.xlsx")
    payload = extract_ir_json(content, format="xlsx")
    file_regression.check(
        _canonical_json(payload),
        extension=".json",
        basename="trim_all_empty_xlsx_ir",
    )


def test_trim_all_empty_xlsx_markdown_snapshot(file_regression) -> None:
    content = _read_fixture("trim_all_empty.xlsx")
    payload = markdown_from_bytes(content, format="xlsx")
    file_regression.check(
        _canonical_markdown(payload),
        extension=".md",
        basename="trim_all_empty_xlsx_markdown",
    )
