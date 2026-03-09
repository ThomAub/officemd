"""Diff utilities for comparing markdown outputs."""

from __future__ import annotations

import difflib

from rich.console import Console
from rich.syntax import Syntax


def diff_markdown(a: str, b: str, context: int = 3) -> str:
    """Return a unified diff between two markdown strings."""
    a_lines = a.splitlines(keepends=True)
    b_lines = b.splitlines(keepends=True)
    diff = difflib.unified_diff(a_lines, b_lines, fromfile="a", tofile="b", n=context)
    return "".join(diff)


def render_diff(a: str, b: str, context: int = 3) -> None:
    """Pretty-print a unified diff to the terminal using rich."""
    diff_text = diff_markdown(a, b, context=context)
    if not diff_text:
        Console().print("[dim]No differences found.[/dim]")
        return

    syntax = Syntax(diff_text, "diff", theme="monokai")
    Console().print(syntax)
