#!/usr/bin/env python3
"""Review insta markdown snapshot changes with Rich.

This script is intentionally focused on `*_markdown.snap` files. It extracts the
snapshot payload after the insta YAML-ish front matter, builds a git-style
unified diff, and renders changed markdown hunks so the review feels closer to
reading the document than reading raw snapshot text.
"""

from __future__ import annotations

import argparse
import difflib
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence

from rich import box
from rich.console import Console, ConsoleOptions, Group, RenderableType, RenderResult
from rich.markdown import ImageItem, Markdown
from rich.panel import Panel
from rich.segment import Segment
from rich.style import Style
from rich.table import Table
from rich.text import Text

REPO_ROOT = Path(__file__).resolve().parent.parent
SNAPSHOT_GLOB = "**/*_markdown.snap"


class SnapshotImageItem(ImageItem):
    """Render markdown images as explicit, review-friendly placeholders."""

    def __rich_console__(
        self, console: Console, options: ConsoleOptions
    ) -> RenderResult:
        title = self.text.plain if self.text else "image"
        title = title.strip() or "image"
        destination = self.destination.strip()
        label = f"🖼  {title}"
        if destination:
            label = f"{label} → {destination}"
        text = Text(label, style="markdown.image")
        if self.hyperlinks and destination:
            text.stylize(Style(link=destination))
        yield text


class SnapshotMarkdown(Markdown):
    """Rich Markdown with image placeholders tuned for snapshot review."""

    elements = {**Markdown.elements, "image": SnapshotImageItem}


@dataclass(frozen=True)
class AnchoredHighlight:
    target: str
    prefix: str
    suffix: str
    style: str


class HighlightedMarkdown:
    """Render markdown, then overlay anchored diff highlights on its segments."""

    def __init__(self, markdown: str, highlights: Sequence[AnchoredHighlight]) -> None:
        self.markdown = markdown
        self.highlights = highlights

    def __rich_console__(
        self, console: Console, options: ConsoleOptions
    ) -> RenderResult:
        rendered = list(
            console.render(
                SnapshotMarkdown(
                    self.markdown,
                    code_theme="ansi_dark",
                    hyperlinks=False,
                ),
                options,
            )
        )
        plain = "".join(segment.text for segment in rendered if not segment.control)
        spans = locate_highlight_spans(plain, self.highlights)
        yield from apply_highlight_spans(console, rendered, spans)


@dataclass(frozen=True)
class SnapshotContent:
    path: Path
    label: str
    text: str
    body: str


@dataclass(frozen=True)
class Hunk:
    header: str
    old_lines: list[str]
    new_lines: list[str]
    diff_lines: list[str]


class ReviewTheme:
    added = "bold green"
    removed = "bold red"
    hunk = "bold cyan"
    file = "bold magenta"
    meta = "dim"
    warning = "bold yellow"


def run_git(args: Sequence[str], *, check: bool = True) -> str:
    completed = subprocess.run(
        ["git", *args],
        cwd=REPO_ROOT,
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if check and completed.returncode != 0:
        message = completed.stderr.strip() or completed.stdout.strip()
        raise RuntimeError(f"git {' '.join(args)} failed: {message}")
    return completed.stdout


def repo_relative(path: Path) -> str:
    try:
        return path.resolve().relative_to(REPO_ROOT).as_posix()
    except ValueError:
        return path.as_posix()


def split_snapshot(text: str) -> str:
    """Return the snapshot payload, dropping insta front matter if present."""
    lines = text.splitlines(keepends=True)
    if lines and lines[0].strip() == "---":
        for index in range(1, len(lines)):
            if lines[index].strip() == "---":
                return "".join(lines[index + 1 :])
    return text


def read_worktree(path: Path, label: str | None = None) -> SnapshotContent:
    text = path.read_text(encoding="utf-8")
    return SnapshotContent(
        path=path,
        label=label or repo_relative(path),
        text=text,
        body=split_snapshot(text),
    )


def read_git_head(path: Path) -> SnapshotContent | None:
    rel = repo_relative(path)
    completed = subprocess.run(
        ["git", "show", f"HEAD:{rel}"],
        cwd=REPO_ROOT,
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if completed.returncode != 0:
        return None
    text = completed.stdout
    return SnapshotContent(
        path=path, label=f"HEAD:{rel}", text=text, body=split_snapshot(text)
    )


def snap_partner(path: Path) -> Path | None:
    name = path.name
    if name.endswith(".snap.new"):
        return path.with_name(name.removesuffix(".new"))
    candidate = path.with_name(f"{name}.new")
    return candidate if candidate.exists() else None


def discover_changed_snapshots() -> list[Path]:
    porcelain = run_git(
        ["status", "--porcelain", "--", "crates/tests/rust_snapshots/tests/snapshots"]
    )
    paths: list[Path] = []
    for line in porcelain.splitlines():
        if not line:
            continue
        raw = line[3:]
        # Porcelain v1 rename lines look like `old -> new`; use the new side.
        if " -> " in raw:
            raw = raw.rsplit(" -> ", 1)[1]
        path = (REPO_ROOT / raw).resolve()
        if is_markdown_snapshot(path):
            paths.append(path)
    return sorted(dict.fromkeys(paths))


def is_markdown_snapshot(path: Path) -> bool:
    return path.name.endswith("_markdown.snap") or path.name.endswith(
        "_markdown.snap.new"
    )


def find_all_markdown_snapshots() -> list[Path]:
    root = REPO_ROOT / "crates" / "tests" / "rust_snapshots" / "tests" / "snapshots"
    return sorted(root.glob(SNAPSHOT_GLOB))


def build_hunks(
    old: SnapshotContent, new: SnapshotContent, *, context: int
) -> list[Hunk]:
    diff = list(
        difflib.unified_diff(
            old.body.splitlines(),
            new.body.splitlines(),
            fromfile=old.label,
            tofile=new.label,
            n=context,
            lineterm="",
        )
    )
    hunks: list[Hunk] = []
    current_header: str | None = None
    current_lines: list[str] = []

    def flush() -> None:
        nonlocal current_header, current_lines
        if current_header is None:
            return
        old_lines: list[str] = []
        new_lines: list[str] = []
        for line in current_lines:
            if line.startswith("-") and not line.startswith("---"):
                old_lines.append(line[1:])
            elif line.startswith("+") and not line.startswith("+++"):
                new_lines.append(line[1:])
            elif line.startswith(" "):
                old_lines.append(line[1:])
                new_lines.append(line[1:])
        hunks.append(
            Hunk(
                header=current_header,
                old_lines=old_lines,
                new_lines=new_lines,
                diff_lines=current_lines,
            )
        )
        current_header = None
        current_lines = []

    for line in diff:
        if line.startswith("@@"):
            flush()
            current_header = line
        elif current_header is not None:
            current_lines.append(line)
    flush()
    return hunks


def line_number_from_hunk_header(header: str) -> tuple[int, int]:
    # Unified diff header shape: @@ -old_start,old_count +new_start,new_count @@
    parts = header.split()
    old_start = int(parts[1].split(",", 1)[0].removeprefix("-"))
    new_start = int(parts[2].split(",", 1)[0].removeprefix("+"))
    return old_start, new_start


def diff_line_table(hunk: Hunk) -> Table:
    old_line, new_line = line_number_from_hunk_header(hunk.header)
    table = Table.grid(expand=True)
    table.add_column("old", justify="right", width=5, style=ReviewTheme.meta)
    table.add_column("new", justify="right", width=5, style=ReviewTheme.meta)
    table.add_column("mark", width=1)
    table.add_column("text", ratio=1, overflow="fold")
    table.add_row("", "", "", Text(hunk.header, style=ReviewTheme.hunk))
    for raw in hunk.diff_lines:
        if raw.startswith("-") and not raw.startswith("---"):
            table.add_row(
                str(old_line), "", "-", Text(raw[1:], style=ReviewTheme.removed)
            )
            old_line += 1
        elif raw.startswith("+") and not raw.startswith("+++"):
            table.add_row(
                "", str(new_line), "+", Text(raw[1:], style=ReviewTheme.added)
            )
            new_line += 1
        elif raw.startswith(" "):
            table.add_row(
                str(old_line), str(new_line), " ", Text(raw[1:], style="default")
            )
            old_line += 1
            new_line += 1
        else:
            table.add_row("", "", "", Text(raw, style=ReviewTheme.meta))
    return table


def changed_line_pairs(hunk: Hunk) -> list[tuple[str, str]]:
    removed: list[str] = []
    added: list[str] = []
    pairs: list[tuple[str, str]] = []

    def flush() -> None:
        nonlocal removed, added
        for old_line, new_line in zip(removed, added, strict=False):
            pairs.append((old_line, new_line))
        removed = []
        added = []

    for raw in hunk.diff_lines:
        if raw.startswith("-") and not raw.startswith("---"):
            removed.append(raw[1:])
        elif raw.startswith("+") and not raw.startswith("+++"):
            added.append(raw[1:])
        else:
            flush()
    flush()
    return pairs


def anchored_diff_highlights(
    hunk: Hunk,
) -> tuple[list[AnchoredHighlight], list[AnchoredHighlight]]:
    before: list[AnchoredHighlight] = []
    after: list[AnchoredHighlight] = []
    matcher_context = 24

    for old_line, new_line in changed_line_pairs(hunk):
        matcher = difflib.SequenceMatcher(None, old_line, new_line, autojunk=False)
        for tag, old_start, old_end, new_start, new_end in matcher.get_opcodes():
            if tag == "equal":
                continue
            old_start, old_end = expand_change_to_word(old_line, old_start, old_end)
            new_start, new_end = expand_change_to_word(new_line, new_start, new_end)
            old_target = old_line[old_start:old_end]
            new_target = new_line[new_start:new_end]
            old_prefix = old_line[max(0, old_start - matcher_context) : old_start]
            old_suffix = old_line[old_end : old_end + matcher_context]
            new_prefix = new_line[max(0, new_start - matcher_context) : new_start]
            new_suffix = new_line[new_end : new_end + matcher_context]
            if old_target:
                before.append(
                    AnchoredHighlight(
                        target=old_target,
                        prefix=old_prefix,
                        suffix=old_suffix,
                        style="bold white on red3",
                    )
                )
            if new_target:
                after.append(
                    AnchoredHighlight(
                        target=new_target,
                        prefix=new_prefix,
                        suffix=new_suffix,
                        style="bold white on green4",
                    )
                )
    return before, after


def expand_change_to_word(line: str, start: int, end: int) -> tuple[int, int]:
    """Expand a character-level edit to a visible word-sized highlight."""
    if not line:
        return start, end

    if start == end:
        if start < len(line) and line[start].isalnum():
            end = start + 1
        elif start > 0 and line[start - 1].isalnum():
            start -= 1
        else:
            return start, end

    while start > 0 and line[start - 1].isalnum():
        start -= 1
    while end < len(line) and line[end].isalnum():
        end += 1
    return start, end


def locate_highlight_spans(
    plain: str, highlights: Sequence[AnchoredHighlight]
) -> list[tuple[int, int, str]]:
    spans: list[tuple[int, int, str]] = []
    search_from = 0
    for highlight in highlights:
        start = locate_highlight(plain, highlight, search_from)
        if start == -1:
            start = locate_highlight(plain, highlight, 0)
        if start == -1:
            continue
        end = start + len(highlight.target)
        spans.append((start, end, highlight.style))
        search_from = end
    return spans


def locate_highlight(plain: str, highlight: AnchoredHighlight, start: int) -> int:
    target_start = plain.find(highlight.target, start)
    while target_start != -1:
        prefix_ok = not highlight.prefix or plain[
            max(0, target_start - len(highlight.prefix)) : target_start
        ].endswith(highlight.prefix)
        target_end = target_start + len(highlight.target)
        suffix_ok = not highlight.suffix or plain[
            target_end : target_end + len(highlight.suffix)
        ].startswith(highlight.suffix)
        if prefix_ok and suffix_ok:
            return target_start
        target_start = plain.find(highlight.target, target_start + 1)
    return -1


def apply_highlight_spans(
    console: Console,
    rendered: Sequence[Segment],
    spans: Sequence[tuple[int, int, str]],
) -> RenderResult:
    if not spans:
        yield from rendered
        return

    position = 0
    span_index = 0
    for segment in rendered:
        if segment.control or not segment.text:
            yield segment
            continue
        text_start = position
        text_end = position + len(segment.text)
        cursor = 0
        while span_index < len(spans) and spans[span_index][1] <= text_start:
            span_index += 1
        local_span_index = span_index
        while local_span_index < len(spans):
            start, end, style_name = spans[local_span_index]
            if start >= text_end:
                break
            if end <= text_start:
                local_span_index += 1
                continue
            overlap_start = max(start, text_start) - text_start
            overlap_end = min(end, text_end) - text_start
            if cursor < overlap_start:
                yield Segment(segment.text[cursor:overlap_start], segment.style)
            highlight_style = (
                segment.style + console.get_style(style_name)
                if segment.style
                else console.get_style(style_name)
            )
            yield Segment(segment.text[overlap_start:overlap_end], highlight_style)
            cursor = overlap_end
            if end <= text_end:
                local_span_index += 1
            else:
                break
        if cursor < len(segment.text):
            yield Segment(segment.text[cursor:], segment.style)
        position = text_end


def markdown_panel(
    title: str,
    lines: Iterable[str],
    border_style: str,
    highlights: Sequence[AnchoredHighlight],
) -> Panel:
    text = "\n".join(lines).strip("\n")
    if not text.strip():
        renderable = Text("∅", style=ReviewTheme.meta)
    else:
        renderable = HighlightedMarkdown(text, highlights)
    return Panel(renderable, title=title, border_style=border_style, box=box.ROUNDED)


def side_by_side_markdown(hunk: Hunk) -> Table:
    before_highlights, after_highlights = anchored_diff_highlights(hunk)
    table = Table.grid(expand=True)
    table.add_column(ratio=1)
    table.add_column(ratio=1)
    table.add_row(
        markdown_panel("before", hunk.old_lines, "red", before_highlights),
        markdown_panel("after", hunk.new_lines, "green", after_highlights),
    )
    return table


def hunk_renderable(hunk: Hunk, *, render_markdown: bool) -> Group:
    parts: list[RenderableType] = [diff_line_table(hunk)]
    if render_markdown:
        parts.extend([Text(""), side_by_side_markdown(hunk)])
    return Group(*parts)


def compare_pair(
    old: SnapshotContent,
    new: SnapshotContent,
    *,
    context: int,
    render_markdown: bool,
) -> Group | None:
    hunks = build_hunks(old, new, context=context)
    if not hunks:
        return None
    title = Text.assemble(
        ("snapshot diff ", ReviewTheme.file),
        (new.label, "bold"),
        (f" ({len(hunks)} hunk{'s' if len(hunks) != 1 else ''})", ReviewTheme.meta),
    )
    parts: list[RenderableType] = [Panel(title, style=ReviewTheme.file, box=box.HEAVY)]
    for index, hunk in enumerate(hunks, start=1):
        parts.append(
            Panel(
                hunk_renderable(hunk, render_markdown=render_markdown),
                title=f"hunk {index}",
                border_style="cyan",
                box=box.SIMPLE_HEAVY,
            )
        )
    return Group(*parts)


def pair_from_path(path: Path) -> tuple[SnapshotContent, SnapshotContent] | None:
    path = path.resolve()
    partner = snap_partner(path)
    if path.name.endswith(".snap.new"):
        old_path = partner
        new_path = path
        if old_path is None or not old_path.exists():
            return None
        return read_worktree(old_path), read_worktree(new_path)
    if partner is not None and partner.exists():
        return read_worktree(path), read_worktree(partner)
    old = read_git_head(path)
    if old is None:
        return None
    return old, read_worktree(path, label=f"worktree:{repo_relative(path)}")


def parse_args(argv: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Render markdown snapshot diffs with Rich, including rendered markdown hunks.",
    )
    parser.add_argument(
        "paths",
        nargs="*",
        type=Path,
        help=(
            "Snapshot paths. With no paths, changed markdown snapshots are discovered from git. "
            "With one path, compare .snap↔.snap.new when present or HEAD↔worktree. "
            "With two paths, compare them directly."
        ),
    )
    parser.add_argument(
        "--context", "-U", type=int, default=4, help="Number of raw diff context lines."
    )
    parser.add_argument(
        "--no-markdown",
        action="store_true",
        help="Only print the raw unified diff table, without rendered markdown previews.",
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="Review all committed markdown snapshots against HEAD/worktree instead of git status changes.",
    )
    parser.add_argument(
        "--width",
        type=int,
        default=None,
        help="Console width override for narrow or wide terminals.",
    )
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    console = Console(width=args.width)
    render_markdown = not args.no_markdown

    try:
        if len(args.paths) == 2:
            pairs = [(read_worktree(args.paths[0]), read_worktree(args.paths[1]))]
        elif len(args.paths) == 1:
            pair = pair_from_path(args.paths[0])
            if pair is None:
                console.print(
                    f"Could not find a comparison source for {args.paths[0]}",
                    style=ReviewTheme.warning,
                )
                return 2
            pairs = [pair]
        elif args.all:
            pairs = [
                pair
                for path in find_all_markdown_snapshots()
                if (pair := pair_from_path(path))
            ]
        else:
            changed = discover_changed_snapshots()
            if not changed:
                console.print(
                    "No changed markdown snapshots found.", style=ReviewTheme.meta
                )
                return 0
            pairs = [pair for path in changed if (pair := pair_from_path(path))]
    except RuntimeError as error:
        console.print(str(error), style="bold red")
        return 2

    rendered_count = 0
    for old, new in pairs:
        renderable = compare_pair(
            old,
            new,
            context=args.context,
            render_markdown=render_markdown,
        )
        if renderable is None:
            continue
        if rendered_count:
            console.print()
        console.print(renderable)
        rendered_count += 1

    if rendered_count == 0:
        console.print("No snapshot diffs to render.", style=ReviewTheme.meta)
    return 1 if rendered_count else 0


if __name__ == "__main__":
    raise SystemExit(main())
