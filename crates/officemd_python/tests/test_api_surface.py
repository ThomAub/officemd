import inspect
from pathlib import Path

import officemd

EXPECTED_EXPORTS = [
    "apply_ooxml_patch_json",
    "create_document_from_markdown",
    "detect_format",
    "docling_from_bytes",
    "extract_csv_tables_ir_json",
    "extract_ir_json",
    "extract_sheet_names",
    "extract_tables_ir_json",
    "inspect_pdf_fonts_json",
    "inspect_pdf_json",
    "markdown_from_bytes",
    "markdown_from_bytes_batch",
    "patch_docx",
    "patch_pptx",
]


def test_binding_surface_exports_exist() -> None:
    for name in EXPECTED_EXPORTS:
        assert callable(getattr(officemd, name))


def test_create_document_from_markdown_returns_ooxml_bytes() -> None:
    content = officemd.create_document_from_markdown("## Section: body\n\nHello\n", "docx")
    assert isinstance(content, bytes)
    assert content[:2] == b"PK"


def test_markdown_from_bytes_signature_includes_force_extract() -> None:
    signature = inspect.signature(officemd.markdown_from_bytes)
    assert list(signature.parameters) == [
        "content",
        "format",
        "include_document_properties",
        "use_first_row_as_header",
        "include_headers_footers",
        "include_formulas",
        "markdown_style",
        "force_extract",
    ]


def test_apply_ooxml_patch_json_returns_edited_ooxml_bytes() -> None:
    content = officemd.create_document_from_markdown("## Section: body\n\nHello\n", "docx")
    patched = officemd.apply_ooxml_patch_json(
        content,
        '{"edits":[{"part":"word/document.xml","from":"Hello","to":"Hello from Python"}]}',
    )
    assert isinstance(patched, bytes)
    markdown = officemd.markdown_from_bytes(patched, format="docx", include_document_properties=True)
    assert "Hello from Python" in markdown


def test_patch_docx_typed_api_replaces_all_text() -> None:
    content = officemd.create_document_from_markdown("## Section: body\n\nword word\n", "docx")
    patched = officemd.patch_docx(
        content,
        officemd.DocxPatch(
            scoped_replacements=[
                officemd.ScopedDocxReplace(
                    officemd.DocxTextScope.ALL_TEXT,
                    officemd.TextReplace(
                        from_text="word",
                        to_text="term",
                        match_policy=officemd.MatchPolicy.WHOLE_WORD,
                    ),
                )
            ],
        ),
    )
    markdown = officemd.markdown_from_bytes(patched, format="docx", include_document_properties=True)
    assert "term term" in markdown


def test_patch_files_can_apply_same_docx_patch_to_multiple_files(tmp_path: Path) -> None:
    content = officemd.create_document_from_markdown("## Section: body\n\nword\n", "docx")
    src1 = tmp_path / "a.docx"
    src2 = tmp_path / "b.docx"
    src1.write_bytes(content)
    src2.write_bytes(content)
    patch = officemd.DocxPatch(
        scoped_replacements=[
            officemd.ScopedDocxReplace(
                officemd.DocxTextScope.ALL_TEXT,
                officemd.TextReplace("word", "term"),
            )
        ]
    )
    results = officemd.patch_files(
        [
            officemd.BatchPatchJob(src1, tmp_path / "out1.docx", patch, "docx"),
            officemd.BatchPatchJob(src2, tmp_path / "out2.docx", patch, "docx"),
        ],
        workers=2,
    )
    assert all(result.ok for result in results)
    assert "term" in officemd.markdown_from_bytes((tmp_path / "out1.docx").read_bytes(), format="docx")
    assert "term" in officemd.markdown_from_bytes((tmp_path / "out2.docx").read_bytes(), format="docx")
