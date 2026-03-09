#!/usr/bin/env python3
"""Compare XLSX markdown output with and without trim_empty.

Usage (after `cd crates/officemd_python && uv run maturin develop --release && cd ../..`):
    uv run --project crates/officemd_python python examples/python/compare_trim_xlsx.py <path.xlsx>
"""

from __future__ import annotations

import sys
from pathlib import Path

from officemd import markdown_from_bytes


def compare(path: Path) -> None:
    content = path.read_bytes()

    md_compact = markdown_from_bytes(content, format="xlsx", markdown_style="compact")
    md_human = markdown_from_bytes(content, format="xlsx", markdown_style="human")

    print(f"=== {path.name} ===\n")

    print(f"--- compact/LlmCompact (trim_empty=true, {len(md_compact)} chars) ---")
    print(md_compact)

    print(f"--- human (trim_empty=false, {len(md_human)} chars) ---")
    print(md_human)

    saved = len(md_human) - len(md_compact)
    pct = (saved / len(md_human) * 100) if md_human else 0.0
    print(f"--- Savings: {saved} chars ({pct:.1f}% reduction) ---\n")


def main() -> None:
    if len(sys.argv) < 2:
        print("Usage: compare_trim_xlsx.py <path.xlsx> [path2.xlsx ...]", file=sys.stderr)
        sys.exit(2)

    for arg in sys.argv[1:]:
        compare(Path(arg))


if __name__ == "__main__":
    main()
