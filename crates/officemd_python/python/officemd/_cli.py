"""CLI entry point for officemd."""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections.abc import Callable
from pathlib import Path

from rich.console import Console

from officemd import detect_format, inspect_pdf_json, markdown_from_bytes

_PAGE_HEADING_RE = re.compile(r"^##\s+Page:\s+(\d+)\s*$")
_SLIDE_HEADING_RE = re.compile(r"^##\s+Slide\s+(\d+)\b.*$")
_SHEET_HEADING_RE = re.compile(r"^##\s+Sheet:\s+(.+?)\s*$")
_RANGE_TOKEN_RE = re.compile(r"^(\d+)\s*-\s*(\d+)$")

_stderr = Console(stderr=True)

_SUPPORTED_FORMATS = ".docx, .xlsx, .csv, .pptx, .pdf"


class CliUsageError(ValueError):
    """Raised for user-facing CLI usage errors."""


def _format_error(exc: Exception, path: Path | None = None) -> str:
    """Return a user-friendly error message for common extraction failures."""
    msg = str(exc)

    if isinstance(exc, FileNotFoundError):
        return f"File not found: {path or msg}"

    if "ZIP error" in msg or "EOCD" in msg or "invalid Zip archive" in msg:
        return (
            "This file appears to be encrypted or not a valid OOXML document. "
            "If the file is password-protected, remove the password and try again."
        )

    if "Could not detect format" in msg:
        return f"Could not detect document format. Supported formats: {_SUPPORTED_FORMATS}"

    return msg


def _warn_scanned_pdf(content: bytes, *, force: bool) -> None:
    """Warn on stderr when a PDF is scanned/image-based."""
    try:
        raw = inspect_pdf_json(content)
        diag = json.loads(raw)
    except Exception:
        return

    classification = diag.get("classification", "")
    if classification not in ("Scanned", "ImageBased"):
        return

    confidence = diag.get("confidence", 0)
    page_count = diag.get("page_count", 0)
    pages_needing_ocr = diag.get("pages_needing_ocr", [])

    if force:
        _stderr.print(
            f"[bold blue]Info:[/bold blue] PDF classified as [bold]{classification}[/bold] "
            f"(confidence: {confidence:.0%}, {page_count} page(s)). "
            "Forced extraction attempted - output may be empty or incomplete."
        )
    else:
        ocr_summary = (
            f"pages needing OCR: {', '.join(str(p) for p in pages_needing_ocr)}"
            if pages_needing_ocr
            else f"{page_count} page(s)"
        )
        _stderr.print(
            f"[bold yellow]Warning:[/bold yellow] PDF classified as [bold]{classification}[/bold] "
            f"(confidence: {confidence:.0%}, {ocr_summary}). "
            "No text could be extracted - this document likely needs OCR.\n"
            "Hint: use [bold]--force[/bold] to attempt extraction anyway."
        )


def _build_render_options(args: argparse.Namespace) -> dict:
    """Build keyword arguments for markdown_from_bytes from parsed CLI args."""
    opts: dict = {}
    if args.format is not None:
        opts["format"] = args.format
    opts["include_document_properties"] = args.include_document_properties
    opts["use_first_row_as_header"] = args.use_first_row_as_header
    opts["include_headers_footers"] = args.include_headers_footers
    opts["markdown_style"] = args.markdown_style
    if getattr(args, "force", False):
        opts["force_extract"] = True
    return opts


def _add_selection_flags(parser: argparse.ArgumentParser) -> None:
    """Add page/sheet selector flags for markdown-oriented subcommands."""
    parser.add_argument(
        "--pages",
        default=None,
        help="Select PDF pages/PPTX slides or XLSX/CSV sheet indices (examples: 1, 1,3-5)",
    )
    parser.add_argument(
        "--sheets",
        default=None,
        help="Select sheets by index/name (examples: 1, 2-4, Sales,Summary)",
    )


def _add_shared_flags(parser: argparse.ArgumentParser) -> None:
    """Add flags shared across subcommands."""
    parser.add_argument(
        "--format",
        choices=["docx", "xlsx", "csv", "pptx", "pdf"],
        default=None,
        help="Force document format instead of auto-detecting",
    )
    parser.add_argument(
        "--include-document-properties",
        action="store_true",
        default=False,
        help="Include document properties in output",
    )
    parser.add_argument(
        "--use-first-row-as-header",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Use first row as table header (default: true)",
    )
    parser.add_argument(
        "--include-headers-footers",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Include headers and footers (default: true)",
    )
    parser.add_argument(
        "--markdown-style",
        choices=["compact", "human"],
        default="compact",
        help="Markdown profile (default: compact)",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        default=False,
        help="Force extraction even for scanned/image-based PDFs",
    )


def _resolve_format(path: Path, explicit_format: str | None, content: bytes) -> str | None:
    """Resolve format from explicit CLI option, path, or content sniffing."""
    if explicit_format is not None:
        return explicit_format
    inferred = _infer_format_from_path(path)
    if inferred is not None:
        return inferred
    try:
        return detect_format(content).removeprefix(".")
    except Exception:
        return None


def _extract_markdown(path: Path, opts: dict) -> tuple[str, str | None]:
    """Read a file and extract markdown plus resolved format."""
    if not path.exists():
        raise FileNotFoundError(path)
    if path.stat().st_size == 0:
        raise ValueError(f"File is empty: {path}")

    content = path.read_bytes()
    merged_opts = dict(opts)
    fmt = _resolve_format(path, merged_opts.get("format"), content)
    if fmt is not None:
        merged_opts["format"] = fmt
    md = markdown_from_bytes(content, **merged_opts)

    # Post-hoc scanned PDF warning
    force = merged_opts.get("force_extract", False)
    if fmt == "pdf" and len(md.strip()) < 50:
        _warn_scanned_pdf(content, force=force)

    return md, fmt


def _infer_format_from_path(path: Path) -> str | None:
    suffix = path.suffix.lower()
    if suffix in {".docx", ".xlsx", ".csv", ".pptx", ".pdf"}:
        return suffix.removeprefix(".")
    return None


def _parse_positive_int(raw: str, *, flag: str) -> int:
    if not raw.isdigit():
        raise CliUsageError(f"Invalid `{flag}` token `{raw}`: expected a positive integer")
    value = int(raw)
    if value < 1:
        raise CliUsageError(f"Invalid `{flag}` token `{raw}`: expected value >= 1")
    return value


def _parse_numeric_selector(selector: str, *, flag: str) -> set[int]:
    """Parse selector grammar like: 1,3-5."""
    selected: set[int] = set()
    if not selector.strip():
        raise CliUsageError(f"`{flag}` cannot be empty")
    for chunk in selector.split(","):
        token = chunk.strip()
        if not token:
            raise CliUsageError(f"Invalid `{flag}` selector `{selector}`: empty token")
        range_match = _RANGE_TOKEN_RE.fullmatch(token)
        if range_match:
            start = _parse_positive_int(range_match.group(1), flag=flag)
            end = _parse_positive_int(range_match.group(2), flag=flag)
            if start > end:
                raise CliUsageError(f"Invalid `{flag}` token `{token}`: start must be <= end")
            selected.update(range(start, end + 1))
            continue
        selected.add(_parse_positive_int(token, flag=flag))
    return selected


def _parse_sheet_selector(selector: str) -> tuple[set[int], set[str]]:
    """Parse sheet selector grammar: indices/ranges and names, comma-separated."""
    indices: set[int] = set()
    names: set[str] = set()
    if not selector.strip():
        raise CliUsageError("`--sheets` cannot be empty")
    for chunk in selector.split(","):
        token = chunk.strip()
        if not token:
            raise CliUsageError(f"Invalid `--sheets` selector `{selector}`: empty token")
        range_match = _RANGE_TOKEN_RE.fullmatch(token)
        if range_match:
            start = _parse_positive_int(range_match.group(1), flag="--sheets")
            end = _parse_positive_int(range_match.group(2), flag="--sheets")
            if start > end:
                raise CliUsageError(f"Invalid `--sheets` token `{token}`: start must be <= end")
            indices.update(range(start, end + 1))
            continue
        if token.isdigit():
            indices.add(_parse_positive_int(token, flag="--sheets"))
            continue
        names.add(token)
    return indices, names


def _line_markers(
    markdown: str, heading_pattern: re.Pattern[str]
) -> list[tuple[int, re.Match[str]]]:
    lines = markdown.splitlines(keepends=True)
    markers: list[tuple[int, re.Match[str]]] = []
    for idx, line in enumerate(lines):
        match = heading_pattern.fullmatch(line.strip())
        if match:
            markers.append((idx, match))
    return markers


def _select_markdown_sections(
    markdown: str,
    *,
    heading_pattern: re.Pattern[str],
    keep_marker: Callable[[re.Match[str], int], bool],
) -> str:
    lines = markdown.splitlines(keepends=True)
    markers = _line_markers(markdown, heading_pattern)
    if not markers:
        return ""

    selected_chunks: list[str] = []
    prefix = "".join(lines[: markers[0][0]])
    selected_any = False

    for i, (start, match) in enumerate(markers):
        if not keep_marker(match, i + 1):
            continue
        if not selected_any and prefix:
            selected_chunks.append(prefix)
        selected_any = True
        end = markers[i + 1][0] if i + 1 < len(markers) else len(lines)
        selected_chunks.append("".join(lines[start:end]))

    if not selected_any:
        return ""
    return "".join(selected_chunks)


def _apply_selectors(
    markdown: str,
    *,
    doc_format: str | None,
    pages_selector: str | None,
    sheets_selector: str | None,
) -> str:
    if pages_selector is None and sheets_selector is None:
        return markdown

    if doc_format in {"pdf", "pptx"}:
        if sheets_selector is not None:
            resolved_format = doc_format or "unknown"
            raise CliUsageError(
                "`--sheets` is only supported for XLSX/CSV input "
                f"(resolved format: {resolved_format})"
            )
        if pages_selector is None:
            return markdown
        selected_pages = _parse_numeric_selector(pages_selector, flag="--pages")
        heading_pattern = _PAGE_HEADING_RE if doc_format == "pdf" else _SLIDE_HEADING_RE
        filtered = _select_markdown_sections(
            markdown,
            heading_pattern=heading_pattern,
            keep_marker=lambda match, _section_index: int(match.group(1)) in selected_pages,
        )
        if not filtered.strip():
            raise CliUsageError(f"`--pages {pages_selector}` did not match any page/slide headings")
        return filtered

    if doc_format in {"xlsx", "csv"}:
        selector_parts = [part for part in (sheets_selector, pages_selector) if part is not None]
        merged_selector = ",".join(selector_parts)
        selected_indices, selected_names = _parse_sheet_selector(merged_selector)
        filtered = _select_markdown_sections(
            markdown,
            heading_pattern=_SHEET_HEADING_RE,
            keep_marker=lambda match, section_index: (
                section_index in selected_indices or match.group(1) in selected_names
            ),
        )
        if not filtered.strip():
            raise CliUsageError(
                f"`--pages/--sheets {merged_selector}` did not match any sheet headings"
            )
        return filtered

    resolved_format = doc_format or "unknown"
    if sheets_selector is not None:
        raise CliUsageError(
            f"`--sheets` is only supported for XLSX/CSV input (resolved format: {resolved_format})"
        )
    raise CliUsageError(
        "`--pages` is supported for PDF/PPTX and XLSX/CSV input "
        f"(resolved format: {resolved_format})"
    )


def cmd_markdown(args: argparse.Namespace) -> None:
    """Extract markdown and print plain text to stdout."""
    opts = _build_render_options(args)
    md, doc_format = _extract_markdown(Path(args.file), opts)
    md = _apply_selectors(
        md,
        doc_format=doc_format,
        pages_selector=getattr(args, "pages", None),
        sheets_selector=getattr(args, "sheets", None),
    )
    sys.stdout.write(md)


def cmd_render(args: argparse.Namespace) -> None:
    """Extract markdown and render to terminal with rich."""
    from officemd.render import render_markdown

    opts = _build_render_options(args)
    md, doc_format = _extract_markdown(Path(args.file), opts)
    md = _apply_selectors(
        md,
        doc_format=doc_format,
        pages_selector=getattr(args, "pages", None),
        sheets_selector=getattr(args, "sheets", None),
    )
    render_markdown(md)


def cmd_diff(args: argparse.Namespace) -> None:
    """Extract markdown from two files and show a diff."""
    from officemd.diff import render_diff

    opts = _build_render_options(args)
    md_a, _ = _extract_markdown(Path(args.file_a), opts)
    md_b, _ = _extract_markdown(Path(args.file_b), opts)
    render_diff(md_a, md_b)


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="officemd",
        description="Extract and render Office documents as markdown",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # markdown subcommand
    p_md = subparsers.add_parser("markdown", help="Extract markdown, print to stdout")
    p_md.add_argument("file", help="Path to an input document")
    _add_shared_flags(p_md)
    _add_selection_flags(p_md)
    p_md.set_defaults(func=cmd_markdown)

    # render subcommand
    p_render = subparsers.add_parser(
        "render", help="Extract markdown, render to terminal with rich"
    )
    p_render.add_argument("file", help="Path to an input document")
    _add_shared_flags(p_render)
    _add_selection_flags(p_render)
    p_render.set_defaults(func=cmd_render)

    # diff subcommand
    p_diff = subparsers.add_parser("diff", help="Diff markdown output of two documents")
    p_diff.add_argument("file_a", help="Path to first input document")
    p_diff.add_argument("file_b", help="Path to second input document")
    _add_shared_flags(p_diff)
    p_diff.set_defaults(func=cmd_diff)

    args = parser.parse_args()
    try:
        args.func(args)
    except CliUsageError as exc:
        parser.error(str(exc))
    except Exception as exc:
        path = getattr(args, "file", None) or getattr(args, "file_a", None)
        path = Path(path) if path else None
        message = _format_error(exc, path)
        _stderr.print(f"[bold red]Error:[/bold red] {message}")
        sys.exit(1)


if __name__ == "__main__":
    main()
