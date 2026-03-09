import json
from pathlib import Path

import pytest


pytest.importorskip("officemd")
from officemd import extract_ir_json  # noqa: E402


def _iter_inlines(section: dict) -> list[dict]:
    inlines: list[dict] = []
    for block in section.get("blocks", []):
        if "Paragraph" in block:
            paragraph = block["Paragraph"]
            inlines.extend(paragraph.get("inlines", []))
    return inlines


def _collect_links(section: dict) -> list[dict]:
    links: list[dict] = []
    for inline in _iter_inlines(section):
        if "Link" in inline:
            links.append(inline["Link"])
    return links


def _has_text(section: dict, needle: str) -> bool:
    for inline in _iter_inlines(section):
        if "Text" in inline and needle in inline["Text"]:
            return True
    return False


def test_extract_ir_json_fixture() -> None:
    fixture = Path(__file__).parent / "fixtures" / "basic.docx"
    if not fixture.exists():
        pytest.skip(f"fixture missing: {fixture}")
    content = fixture.read_bytes()
    doc = json.loads(extract_ir_json(content))

    assert doc["kind"] == "Docx"

    sections = {section["name"]: section for section in doc["sections"]}
    assert "body" in sections
    assert "footnotes" in sections

    body = sections["body"]
    links = _collect_links(body)

    assert any(link.get("target") == "https://example.com" for link in links)
    assert any(link.get("target") == "https://field.example" for link in links)
    assert _has_text(body, "[^c0]")

    comments = body.get("comments", [])
    assert len(comments) == 1
    assert comments[0]["id"] == "c0"
    assert comments[0]["author"] == "Alice"
    assert "Note one" in comments[0]["text"]

    footnotes = sections["footnotes"]
    assert _has_text(footnotes, "Footnote text")
