from pathlib import Path

import pytest

from officemd import markdown_from_bytes


FIXTURE_PATH = Path(__file__).parent / "fixtures" / "sample.xlsx"


def test_markdown_basic_table_and_formula():
    if not FIXTURE_PATH.exists():
        pytest.skip(f"fixture missing: {FIXTURE_PATH}")
    content = FIXTURE_PATH.read_bytes()
    md = markdown_from_bytes(content, use_first_row_as_header=False)

    assert "## Sheet: Sample Sheet" in md
    assert "| Col1 | Col2 | Col3 |" in md
    assert "C2=`=SUM(1,2,3)`" in md
