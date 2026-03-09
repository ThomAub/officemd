"""Terminal markdown rendering using rich."""

from __future__ import annotations

from rich.console import Console
from rich.markdown import Markdown


def render_markdown(markdown: str) -> None:
    """Pretty-print markdown to the terminal using rich."""
    Console().print(Markdown(markdown))
