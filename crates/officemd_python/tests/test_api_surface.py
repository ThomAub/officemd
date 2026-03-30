import inspect
import zipfile
from io import BytesIO
from pathlib import Path

import officemd


def _build_zip(parts: list[tuple[str, str]]) -> bytes:
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, content in parts:
            zf.writestr(name, content)
    return buffer.getvalue()


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
    "patch_docx_batch",
    "patch_docx_batch_with_report",
    "patch_docx_with_report",
    "patch_pptx",
    "patch_pptx_batch",
    "patch_pptx_batch_with_report",
    "patch_pptx_with_report",
    "patch_xlsx",
    "patch_xlsx_batch",
    "patch_xlsx_batch_with_report",
    "patch_xlsx_with_report",
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
    markdown = officemd.markdown_from_bytes(
        patched, format="docx", include_document_properties=True
    )
    assert "Hello from Python" in markdown


def test_patch_docx_typed_api_replaces_all_text() -> None:
    content = _build_zip(
        [
            ("word/document.xml", "<w:t>word word</w:t>"),
            ("docProps/app.xml", "<Properties><Company>word company</Company></Properties>"),
            (
                "docProps/custom.xml",
                '<Properties><property name="FilePath"><vt:lpwstr>word.docx</vt:lpwstr></property></Properties>',
            ),
            (
                "docProps/core.xml",
                '<cp:coreProperties xmlns:cp="cp" xmlns:dc="dc"><dc:title>word title</dc:title></cp:coreProperties>',
            ),
        ]
    )
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
    with zipfile.ZipFile(BytesIO(patched)) as zf:
        assert "term term" in zf.read("word/document.xml").decode()
        assert "term company" in zf.read("docProps/app.xml").decode()
        assert "term.docx" in zf.read("docProps/custom.xml").decode()
        assert "term title" in zf.read("docProps/core.xml").decode()


def test_patch_docx_with_report_returns_counts() -> None:
    content = officemd.create_document_from_markdown("## Section: body\n\nword word\n", "docx")
    patched = officemd.patch_docx_with_report(
        content,
        officemd.DocxPatch(
            scoped_replacements=[
                officemd.ScopedDocxReplace(
                    officemd.DocxTextScope.ALL_TEXT,
                    officemd.TextReplace("word", "term"),
                )
            ],
        ),
    )
    assert patched.report.replacements_applied >= 2
    assert patched.report.parts_modified >= 1
    assert "term term" in officemd.markdown_from_bytes(patched.content, format="docx")


def test_patch_docx_batch_applies_same_patch_with_rayon_workers() -> None:
    content = officemd.create_document_from_markdown("## Section: body\n\nword\n", "docx")
    patch = officemd.DocxPatch(
        scoped_replacements=[
            officemd.ScopedDocxReplace(
                officemd.DocxTextScope.ALL_TEXT,
                officemd.TextReplace("word", "term"),
            )
        ]
    )
    patched = officemd.patch_docx_batch([content, content], patch, workers=2)
    assert len(patched) == 2
    assert "term" in officemd.markdown_from_bytes(patched[0], format="docx")
    assert "term" in officemd.markdown_from_bytes(patched[1], format="docx")


def test_patch_docx_batch_with_report_returns_counts() -> None:
    content = officemd.create_document_from_markdown("## Section: body\n\nword word\n", "docx")
    patch = officemd.DocxPatch(
        scoped_replacements=[
            officemd.ScopedDocxReplace(
                officemd.DocxTextScope.ALL_TEXT,
                officemd.TextReplace("word", "term"),
            )
        ]
    )
    patched = officemd.patch_docx_batch_with_report([content], patch, workers=2)
    assert len(patched) == 1
    assert patched[0].report.replacements_applied >= 2
    assert patched[0].report.parts_modified >= 1
    assert "term term" in officemd.markdown_from_bytes(patched[0].content, format="docx")


def test_patch_docx_can_replace_comment_author_and_metadata_fields() -> None:
    content = _build_zip(
        [
            (
                "word/comments.xml",
                '<w:comments><w:comment w:id="0" w:author="Alice"><w:p><w:r><w:t>Needs review</w:t></w:r></w:p></w:comment></w:comments>',
            ),
            (
                "docProps/app.xml",
                "<Properties><Company>Old Company</Company><Template>Old Template</Template></Properties>",
            ),
            (
                "docProps/custom.xml",
                '<Properties><property name="FilePath"><vt:lpwstr>/tmp/old.docx</vt:lpwstr></property></Properties>',
            ),
            (
                "docProps/core.xml",
                '<cp:coreProperties xmlns:cp="cp" xmlns:dc="dc"><dc:title>old</dc:title></cp:coreProperties>',
            ),
        ]
    )
    patched = officemd.patch_docx(
        content,
        officemd.DocxPatch(
            scoped_replacements=[
                officemd.ScopedDocxReplace(
                    officemd.DocxTextScope.COMMENTS,
                    officemd.TextReplace("Alice", "Bob"),
                ),
                officemd.ScopedDocxReplace(
                    officemd.DocxTextScope.METADATA_APP,
                    officemd.TextReplace("Old", "New"),
                ),
                officemd.ScopedDocxReplace(
                    officemd.DocxTextScope.METADATA_CUSTOM,
                    officemd.TextReplace("/tmp/old.docx", "/tmp/new.docx"),
                ),
            ]
        ),
    )
    with zipfile.ZipFile(BytesIO(patched)) as zf:
        assert 'w:author="Bob"' in zf.read("word/comments.xml").decode()
        assert "New Company" in zf.read("docProps/app.xml").decode()
        assert "/tmp/new.docx" in zf.read("docProps/custom.xml").decode()


def test_patch_pptx_all_text_includes_comment_authors_and_metadata() -> None:
    content = _build_zip(
        [
            ("ppt/slides/slide1.xml", "<a:t>word slide</a:t>"),
            (
                "ppt/commentAuthors.xml",
                '<p:cmAuthorLst><p:cmAuthor id="0" name="word author" initials="WA"/></p:cmAuthorLst>',
            ),
            (
                "docProps/app.xml",
                "<Properties><Company>word company</Company></Properties>",
            ),
            (
                "docProps/custom.xml",
                '<Properties><property name="FileName"><vt:lpwstr>word.pptx</vt:lpwstr></property></Properties>',
            ),
            (
                "docProps/core.xml",
                '<cp:coreProperties xmlns:cp="cp" xmlns:dc="dc"><dc:title>word title</dc:title></cp:coreProperties>',
            ),
        ]
    )
    patched = officemd.patch_pptx(
        content,
        officemd.PptxPatch(
            scoped_replacements=[
                officemd.ScopedPptxReplace(
                    officemd.PptxTextScope.ALL_TEXT,
                    officemd.TextReplace("word", "term"),
                ),
            ]
        ),
    )
    with zipfile.ZipFile(BytesIO(patched)) as zf:
        assert "term slide" in zf.read("ppt/slides/slide1.xml").decode()
        assert 'name="term author"' in zf.read("ppt/commentAuthors.xml").decode()
        assert "term company" in zf.read("docProps/app.xml").decode()
        assert "term.pptx" in zf.read("docProps/custom.xml").decode()
        assert "term title" in zf.read("docProps/core.xml").decode()


def test_patch_pptx_can_replace_comment_author_and_metadata_fields() -> None:
    content = _build_zip(
        [
            (
                "ppt/commentAuthors.xml",
                '<p:cmAuthorLst><p:cmAuthor id="0" name="Alice" initials="AL"/></p:cmAuthorLst>',
            ),
            (
                "docProps/app.xml",
                "<Properties><Company>Old Company</Company><PresentationFormat>Old Deck</PresentationFormat></Properties>",
            ),
            (
                "docProps/custom.xml",
                '<Properties><property name="FileName"><vt:lpwstr>old.pptx</vt:lpwstr></property></Properties>',
            ),
            (
                "docProps/core.xml",
                '<cp:coreProperties xmlns:cp="cp" xmlns:dc="dc"><dc:title>old</dc:title></cp:coreProperties>',
            ),
        ]
    )
    patched = officemd.patch_pptx(
        content,
        officemd.PptxPatch(
            scoped_replacements=[
                officemd.ScopedPptxReplace(
                    officemd.PptxTextScope.COMMENT_AUTHORS,
                    officemd.TextReplace("Alice", "Bob"),
                ),
                officemd.ScopedPptxReplace(
                    officemd.PptxTextScope.METADATA_APP,
                    officemd.TextReplace("Old", "New"),
                ),
                officemd.ScopedPptxReplace(
                    officemd.PptxTextScope.METADATA_CUSTOM,
                    officemd.TextReplace("old.pptx", "new.pptx"),
                ),
            ]
        ),
    )
    with zipfile.ZipFile(BytesIO(patched)) as zf:
        assert 'name="Bob"' in zf.read("ppt/commentAuthors.xml").decode()
        assert "New Company" in zf.read("docProps/app.xml").decode()
        assert "new.pptx" in zf.read("docProps/custom.xml").decode()


def test_patch_xlsx_with_reference_aware_sheet_rename() -> None:
    content = officemd.create_document_from_markdown(
        "## Sheet: Sales\n\n| Item | Value |\n| --- | --- |\n| Revenue | 10 |\n\n"
        "## Sheet: Summary\n\n| Metric | Value |\n| --- | --- |\n| Revenue | 10 |\n\n"
        "B2=`='Sales'!B2`\n",
        "xlsx",
    )
    patched = officemd.patch_xlsx_with_report(
        content,
        officemd.XlsxPatch(
            rename_sheets=[officemd.XlsxSheetRename("Sales", "Revenue")],
        ),
    )
    assert officemd.extract_sheet_names(patched.content) == ["Revenue", "Summary"]
    assert patched.report.replacements_applied >= 2
    with zipfile.ZipFile(BytesIO(patched.content)) as zf:
        workbook_xml = zf.read("xl/workbook.xml").decode()
        sheet_xml = zf.read("xl/worksheets/sheet2.xml").decode()
    assert 'name="Revenue"' in workbook_xml
    assert "Revenue!B2" in sheet_xml or "'Revenue'!B2" in sheet_xml


def test_patch_xlsx_replaces_workbook_text() -> None:
    content = officemd.create_document_from_markdown(
        "## Sheet: Sales\n\n| Term | Value |\n| --- | --- |\n| word | word |\n",
        "xlsx",
    )
    patched = officemd.patch_xlsx(
        content,
        officemd.XlsxPatch(
            scoped_replacements=[
                officemd.ScopedXlsxReplace(
                    officemd.XlsxTextScope.ALL_TEXT,
                    officemd.TextReplace("word", "term"),
                )
            ]
        ),
    )
    assert "term" in officemd.markdown_from_bytes(patched, format="xlsx")


def test_patch_xlsx_all_text_includes_comments_and_metadata() -> None:
    content = _build_zip(
        [
            ("xl/sharedStrings.xml", "<sst><si><t>word cell</t></si></sst>"),
            (
                "xl/comments1.xml",
                '<comments><authors><author>Alice</author></authors><commentList><comment ref="A1" authorId="0"><text><r><t>word comment</t></r></text></comment></commentList></comments>',
            ),
            (
                "xl/persons/person.xml",
                '<personList><person displayName="Alice" id="{1}" userId="word@example.com" providerId="word"/></personList>',
            ),
            ("docProps/app.xml", "<Properties><Company>word company</Company></Properties>"),
            (
                "docProps/custom.xml",
                '<Properties><property name="FilePath"><vt:lpwstr>word.xlsx</vt:lpwstr></property></Properties>',
            ),
            (
                "docProps/core.xml",
                '<cp:coreProperties xmlns:cp="cp" xmlns:dc="dc"><dc:title>word title</dc:title></cp:coreProperties>',
            ),
            (
                "xl/workbook.xml",
                '<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
            ),
        ]
    )
    patched = officemd.patch_xlsx(
        content,
        officemd.XlsxPatch(
            scoped_replacements=[
                officemd.ScopedXlsxReplace(
                    officemd.XlsxTextScope.ALL_TEXT,
                    officemd.TextReplace("word", "term"),
                )
            ]
        ),
    )
    with zipfile.ZipFile(BytesIO(patched)) as zf:
        assert "term cell" in zf.read("xl/sharedStrings.xml").decode()
        assert "term comment" in zf.read("xl/comments1.xml").decode()
        assert "term@example.com" in zf.read("xl/persons/person.xml").decode()
        assert "term company" in zf.read("docProps/app.xml").decode()
        assert "term.xlsx" in zf.read("docProps/custom.xml").decode()
        assert "term title" in zf.read("docProps/core.xml").decode()


def test_patch_files_uses_rust_batch_patching(tmp_path: Path) -> None:
    content = officemd.create_document_from_markdown("## Section: body\n\nword\n", "docx")
    xlsx = officemd.create_document_from_markdown(
        "## Sheet: Sales\n\n| Term | Value |\n| --- | --- |\n| word | 1 |\n",
        "xlsx",
    )
    src1 = tmp_path / "a.docx"
    src2 = tmp_path / "b.docx"
    src3 = tmp_path / "c.xlsx"
    src1.write_bytes(content)
    src2.write_bytes(content)
    src3.write_bytes(xlsx)
    patch = officemd.DocxPatch(
        scoped_replacements=[
            officemd.ScopedDocxReplace(
                officemd.DocxTextScope.ALL_TEXT,
                officemd.TextReplace("word", "term"),
            )
        ]
    )
    xlsx_patch = officemd.XlsxPatch(
        rename_sheets=[officemd.XlsxSheetRename("Sales", "Revenue")],
        scoped_replacements=[
            officemd.ScopedXlsxReplace(
                officemd.XlsxTextScope.ALL_TEXT,
                officemd.TextReplace("word", "term"),
            )
        ],
    )
    results = officemd.patch_files(
        [
            officemd.BatchPatchJob(src1, tmp_path / "out1.docx", patch, "docx"),
            officemd.BatchPatchJob(src2, tmp_path / "out2.docx", patch, "docx"),
            officemd.BatchPatchJob(src3, tmp_path / "out3.xlsx", xlsx_patch, "xlsx"),
        ],
        workers=2,
    )
    assert all(result.ok for result in results)
    assert all(result.report is not None for result in results)
    assert all(
        result.report.replacements_applied >= 1 for result in results if result.report is not None
    )
    assert "term" in officemd.markdown_from_bytes(
        (tmp_path / "out1.docx").read_bytes(), format="docx"
    )
    assert "term" in officemd.markdown_from_bytes(
        (tmp_path / "out2.docx").read_bytes(), format="docx"
    )
    assert "## Sheet: Revenue" in officemd.markdown_from_bytes(
        (tmp_path / "out3.xlsx").read_bytes(), format="xlsx"
    )
