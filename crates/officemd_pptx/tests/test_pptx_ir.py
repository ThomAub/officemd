import json
from pathlib import Path

import pytest


pytest.importorskip("officemd")
from officemd import extract_ir_json  # noqa: E402

FIXTURES_DIR = Path(__file__).parent / "fixtures"
DATA_DIRS = [
    Path(__file__).parent / "data",
    Path(__file__).resolve().parents[3] / "tests" / "data",
]


def test_extract_ir_json_sample_fixture() -> None:
    fixture = FIXTURES_DIR / "sample.pptx"
    if not fixture.exists():
        pytest.skip(f"fixture missing: {fixture}")
    content = fixture.read_bytes()
    payload = json.loads(extract_ir_json(content))

    assert payload["kind"] == "Pptx"
    assert payload["slides"]

    slide = payload["slides"][0]
    assert slide["number"] == 1
    assert slide["title"] == "Welcome"
    assert slide["comments"][0]["author"] == "Alice"
    assert slide["comments"][0]["text"] == "Needs review"
    assert slide["notes"]


def test_extract_ir_json_real_world() -> None:
    candidates = []
    for data_dir in DATA_DIRS:
        if not data_dir.exists():
            continue
        candidates.extend(sorted(data_dir.glob("*.pptx")))
        candidates.extend(sorted(data_dir.glob("*.PPTX")))
    if not candidates:
        pytest.skip("No .pptx files found in tests/data or crates/officemd_pptx/tests/data")

    content = candidates[0].read_bytes()
    payload = json.loads(extract_ir_json(content))
    assert payload["kind"] == "Pptx"
    assert payload["slides"]
