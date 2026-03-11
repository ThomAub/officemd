# /// script
# requires-python = ">=3.12"
# dependencies = ["docling==2.78.0"]
# ///
"""Benchmark docling markdown extraction.

    uv run crates/tests/benchmarks/bench_docling.py examples/data/showcase.docx
"""

import sys
from pathlib import Path

from docling.document_converter import DocumentConverter


def main() -> None:
    if len(sys.argv) < 2:
        print(f"Usage: {sys.argv[0]} <file> [file ...]", file=sys.stderr)
        sys.exit(1)

    converter = DocumentConverter()
    for arg in sys.argv[1:]:
        path = Path(arg)
        if not path.exists():
            print(f"file not found: {path}", file=sys.stderr)
            sys.exit(1)
        result = converter.convert(str(path))
        sys.stdout.write(result.document.export_to_markdown())


if __name__ == "__main__":
    main()
