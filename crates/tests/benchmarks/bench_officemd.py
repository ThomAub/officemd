#!/usr/bin/env python3
"""Benchmark officemd markdown extraction.

Run from the officemd_python directory after `uv run maturin develop --release`:
    uv run python ../../crates/tests/benchmarks/bench_officemd.py ../../examples/data/showcase.docx
"""

import sys
from pathlib import Path

from officemd import markdown_from_bytes

EXT_TO_FORMAT = {".docx": "docx", ".xlsx": "xlsx", ".csv": "csv", ".pptx": "pptx", ".pdf": "pdf"}


def main() -> None:
    if len(sys.argv) < 2:
        print(f"Usage: {sys.argv[0]} <file> [file ...]", file=sys.stderr)
        sys.exit(1)

    for arg in sys.argv[1:]:
        path = Path(arg)
        fmt = EXT_TO_FORMAT.get(path.suffix.lower())
        if fmt is None:
            print(f"unsupported format: {path.suffix}", file=sys.stderr)
            sys.exit(1)
        content = path.read_bytes()
        md = markdown_from_bytes(content, format=fmt)
        sys.stdout.write(md)


if __name__ == "__main__":
    main()
