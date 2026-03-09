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
