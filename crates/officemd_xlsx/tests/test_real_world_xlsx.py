import json
from pathlib import Path

import pytest


pytest.importorskip("officemd")
from officemd import extract_sheet_names, extract_tables_ir_json  # noqa: E402


def _real_world_xlsx_path() -> Path:
    test_dir = Path(__file__).parent
    candidate_dirs = [
        test_dir / "data",
        test_dir.parents[3] / "tests" / "data",
    ]

    candidates: list[Path] = []
    for data_dir in candidate_dirs:
        if not data_dir.exists():
            continue
        candidates.extend(sorted(data_dir.glob("*.xlsx")))
        candidates.extend(sorted(data_dir.glob("*.XLSX")))

    if not candidates:
        pytest.skip("No .xlsx files found in tests/data or crates/officemd_xlsx/tests/data")

    return candidates[0]


def test_real_world_xlsx_extracts_ir_and_markdown() -> None:
    path = _real_world_xlsx_path()
    content = path.read_bytes()

    sheet_names = extract_sheet_names(content)
    assert sheet_names

    ir_payload = json.loads(extract_tables_ir_json(content))
    assert ir_payload["kind"] == "Xlsx"
    assert ir_payload["sheets"]
