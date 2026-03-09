from pathlib import Path

import pytest


pytest.importorskip("officemd")
from officemd import extract_ir_json  # noqa: E402


def _data_files() -> list[Path]:
    test_dir = Path(__file__).parent
    candidate_dirs = [
        test_dir / "data",
        test_dir.parents[3] / "tests" / "data",
    ]
    candidates: list[Path] = []
    for data_dir in candidate_dirs:
        if not data_dir.exists():
            continue
        candidates.extend(sorted(data_dir.glob("*.docx")))
        candidates.extend(sorted(data_dir.glob("*.DOCX")))
    return candidates


def test_real_world_docx_extracts() -> None:
    files = _data_files()
    if not files:
        pytest.skip("No real-world DOCX files found in tests/data or crates/officemd_docx/tests/data")

    content = files[0].read_bytes()
    doc_json = extract_ir_json(content)

    assert "Docx" in doc_json
    assert "sections" in doc_json
