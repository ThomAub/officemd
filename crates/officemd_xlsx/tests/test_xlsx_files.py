import json
from pathlib import Path

import pytest


pytest.importorskip("officemd")
from officemd import extract_ir_json, extract_sheet_names, extract_tables_ir_json, markdown_from_bytes  # noqa: E402

FIXTURES_DIR = Path(__file__).parent / "fixtures"
DATA_DIRS = [
    Path(__file__).parent / "data",
    Path(__file__).resolve().parents[3] / "tests" / "data",
]


def test_extract_sheet_names_sample_fixture() -> None:
    fixture = FIXTURES_DIR / "sample.xlsx"
    if not fixture.exists():
        pytest.skip(f"fixture missing: {fixture}")
    content = fixture.read_bytes()
    sheet_names = extract_sheet_names(content)
    assert sheet_names == ["Sample Sheet"]

    ir_payload = json.loads(extract_ir_json(content))
    assert ir_payload["kind"] == "Xlsx"
    assert ir_payload["sheets"][0]["name"] == "Sample Sheet"


def test_real_world_xlsx_extracts_ir_and_markdown() -> None:
    candidates = []
    for data_dir in DATA_DIRS:
        if not data_dir.exists():
            continue
        candidates.extend(sorted(data_dir.glob("*.xlsx")))
        candidates.extend(sorted(data_dir.glob("*.XLSX")))
    if not candidates:
        pytest.skip("No .xlsx files found in tests/data or crates/officemd_xlsx/tests/data")

    content = candidates[0].read_bytes()

    sheet_names = extract_sheet_names(content)
    assert sheet_names

    ir_payload = json.loads(extract_tables_ir_json(content))
    assert ir_payload["kind"] == "Xlsx"
    assert ir_payload["sheets"]

    markdown = markdown_from_bytes(content)
    assert "## Sheet:" in markdown
