# /// script
# requires-python = ">=3.12"
# dependencies = ["markitdown[all]==0.1.5"]
# ///
"""Benchmark markitdown markdown extraction.

    uv run crates/tests/benchmarks/bench_markitdown.py examples/data/showcase.docx
"""

import sys
from pathlib import Path

from markitdown import MarkItDown


def main() -> None:
    if len(sys.argv) < 2:
        print(f"Usage: {sys.argv[0]} <file> [file ...]", file=sys.stderr)
        sys.exit(1)

    converter = MarkItDown()
    for arg in sys.argv[1:]:
        path = Path(arg)
        if not path.exists():
            print(f"file not found: {path}", file=sys.stderr)
            sys.exit(1)
        result = converter.convert(str(path))
        sys.stdout.write(result.text_content)


if __name__ == "__main__":
    main()
