from pathlib import Path

import pytest


pytest.importorskip("officemd")
from officemd import markdown_from_bytes  # noqa: E402


def test_markdown_from_bytes_fixture() -> None:
    fixture = Path(__file__).parent / "fixtures" / "basic.docx"
    if not fixture.exists():
        pytest.skip(f"fixture missing: {fixture}")
    content = fixture.read_bytes()
    markdown = markdown_from_bytes(content)

    assert "## Section: body" in markdown
    assert "[Example](https://example.com)" in markdown
    assert "[FieldLink](https://field.example)" in markdown
    assert "[^c0]" in markdown
    assert "[^c0]: Alice: Note one" in markdown
    assert "## Section: footnotes" in markdown
    assert "Footnote text" in markdown
