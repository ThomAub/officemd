from __future__ import annotations

import json
import shutil
import subprocess
import tempfile
from pathlib import Path

from officemd import (
    DocxPatch,
    DocxTextScope,
    MatchPolicy,
    PptxPatch,
    PptxTextScope,
    ScopedDocxReplace,
    ScopedPptxReplace,
    ScopedXlsxReplace,
    TextReplace,
    XlsxPatch,
    XlsxSheetRename,
    XlsxTextScope,
    extract_ir_json,
    extract_sheet_names,
    markdown_from_bytes,
    patch_docx,
    patch_pptx,
    patch_xlsx,
    patch_xlsx_with_report,
)

REPO_ROOT = Path(__file__).resolve().parents[2]
DATA_DIR = REPO_ROOT / "examples" / "data"
OUT_DIR = REPO_ROOT / "examples" / "out"


def _check_with_libreoffice(path: Path) -> tuple[bool, str]:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return False, "LibreOffice CLI not found"

    with tempfile.TemporaryDirectory(prefix="officemd-lo-profile-") as profile_dir, tempfile.TemporaryDirectory(
        prefix="officemd-lo-out-"
    ) as out_dir:
        cmd = [
            soffice,
            f"-env:UserInstallation=file://{Path(profile_dir).resolve()}",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            out_dir,
            str(path),
        ]
        completed = subprocess.run(cmd, capture_output=True, text=True, check=False)
        produced = sorted(Path(out_dir).glob("*.pdf"))
        output = (completed.stdout + completed.stderr).strip()
        return completed.returncode == 0 and bool(produced), output or "LibreOffice conversion failed"


def _docx_patch() -> DocxPatch:
    return DocxPatch(
        set_core_title="Edited DOCX Showcase From Python",
        scoped_replacements=[
            ScopedDocxReplace(
                DocxTextScope.HEADERS,
                TextReplace("OOXML Showcase Header", "OfficeMD Showcase Header — edited from Python"),
            ),
            ScopedDocxReplace(
                DocxTextScope.BODY,
                TextReplace("Quarterly Operations Summary", "Quarterly Operations Summary — edited from Python"),
            ),
            ScopedDocxReplace(
                DocxTextScope.COMMENTS,
                TextReplace(
                    "Example DOCX comment captured as markdown footnote.",
                    "Edited DOCX comment from Python patch API.",
                ),
            ),
            ScopedDocxReplace(
                DocxTextScope.METADATA_APP,
                TextReplace("OfficeMD", "OfficeMD Python Example"),
            ),
            ScopedDocxReplace(
                DocxTextScope.METADATA_CUSTOM,
                TextReplace("showcase", "showcase-python"),
            ),
        ],
    )


def _pptx_patch() -> PptxPatch:
    return PptxPatch(
        set_core_title="Edited PPTX Showcase From Python",
        scoped_replacements=[
            ScopedPptxReplace(
                PptxTextScope.ALL_TEXT,
                TextReplace(
                    "Quarterly Review",
                    "Quarterly Review — edited from Python",
                    match_policy=MatchPolicy.EXACT,
                ),
            ),
            ScopedPptxReplace(
                PptxTextScope.COMMENTS,
                TextReplace(
                    "Add one slide on operating margin.",
                    "Edited PPTX comment from Python patch API.",
                ),
            ),
            ScopedPptxReplace(
                PptxTextScope.COMMENT_AUTHORS,
                TextReplace("Alice", "Python Reviewer"),
            ),
            ScopedPptxReplace(
                PptxTextScope.METADATA_APP,
                TextReplace("OfficeMD", "OfficeMD Python Example"),
            ),
        ],
    )


def _xlsx_patch() -> XlsxPatch:
    return XlsxPatch(
        set_core_title="Edited XLSX Showcase From Python",
        rename_sheets=[XlsxSheetRename("Sales", "Sales Python")],
        scoped_replacements=[
            ScopedXlsxReplace(
                XlsxTextScope.ALL_TEXT,
                TextReplace("North", "North Python"),
            ),
            ScopedXlsxReplace(
                XlsxTextScope.COMMENTS,
                TextReplace("Review", "Reviewed from Python"),
            ),
            ScopedXlsxReplace(
                XlsxTextScope.METADATA_APP,
                TextReplace("OfficeMD", "OfficeMD Python Example"),
            ),
            ScopedXlsxReplace(
                XlsxTextScope.METADATA_CUSTOM,
                TextReplace("showcase", "showcase-python"),
            ),
        ],
    )


def _summarize(path: Path, fmt: str) -> dict[str, object]:
    content = path.read_bytes()
    markdown = markdown_from_bytes(content, format=fmt, include_document_properties=True)
    ir = json.loads(extract_ir_json(content, format=fmt))
    libreoffice_ok, libreoffice_output = _check_with_libreoffice(path)
    result: dict[str, object] = {
        "path": str(path.relative_to(REPO_ROOT)),
        "markdown_excerpt": "\n".join(markdown.splitlines()[:18]),
        "libreoffice_ok": libreoffice_ok,
        "libreoffice_output": libreoffice_output,
    }
    if fmt == "docx":
        result["core_title"] = ir["properties"]["core"].get("title")
        result["body_title"] = ir["sections"][0]["blocks"][0]["Paragraph"]["inlines"][0]["Text"]
        result["header_text"] = ir["sections"][1]["blocks"][0]["Paragraph"]["inlines"][0]["Text"]
        result["has_comment"] = "Edited DOCX comment from Python patch API." in markdown
    elif fmt == "pptx":
        result["core_title"] = (ir.get("properties") or {}).get("core", {}).get("title")
        result["slide_1_title"] = ir["slides"][0]["title"]
        result["has_comment"] = "Edited PPTX comment from Python patch API." in markdown
    else:
        result["core_title"] = (ir.get("properties") or {}).get("core", {}).get("title")
        result["sheet_names"] = extract_sheet_names(content)
        result["has_python_edit"] = "North Python" in markdown or "Reviewed from Python" in markdown
    return result


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    print("Plan:")
    print("1. Load examples/data/showcase.docx, showcase.pptx, and showcase.xlsx")
    print("2. Patch OOXML parts directly via officemd.patch_docx / patch_pptx / patch_xlsx")
    print("3. Showcase metadata/comment-author scopes plus ALL_TEXT semantics")
    print("4. Verify edited files via officemd markdown/IR extraction")
    print("5. Verify LibreOffice can still open them via headless PDF conversion")
    print()

    docx_out = OUT_DIR / "showcase_edited.docx"
    docx_out.write_bytes(patch_docx((DATA_DIR / "showcase.docx").read_bytes(), _docx_patch()))

    pptx_out = OUT_DIR / "showcase_edited.pptx"
    pptx_out.write_bytes(patch_pptx((DATA_DIR / "showcase.pptx").read_bytes(), _pptx_patch()))

    xlsx_out = OUT_DIR / "showcase_edited.xlsx"
    xlsx_result = patch_xlsx_with_report((DATA_DIR / "showcase.xlsx").read_bytes(), _xlsx_patch())
    xlsx_out.write_bytes(xlsx_result.content)

    print("DOCX result:")
    print(json.dumps(_summarize(docx_out, "docx"), indent=2, ensure_ascii=False))
    print()
    print("PPTX result:")
    print(json.dumps(_summarize(pptx_out, "pptx"), indent=2, ensure_ascii=False))
    print()
    print("XLSX result:")
    xlsx_summary = _summarize(xlsx_out, "xlsx")
    xlsx_summary["report"] = {
        "parts_scanned": xlsx_result.report.parts_scanned,
        "parts_modified": xlsx_result.report.parts_modified,
        "replacements_applied": xlsx_result.report.replacements_applied,
    }
    print(json.dumps(xlsx_summary, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
