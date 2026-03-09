#!/usr/bin/env python3
import argparse
import json
from pathlib import Path

from officemd import extract_ir_json


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Print the first N PDF pages as markdown from the shared IR output."
    )
    parser.add_argument(
        "pdf_path",
        type=Path,
        nargs="?",
        default=Path("examples/data/OpenXML_WhitePaper.pdf"),
        help="Path to .pdf file (default: examples/data/OpenXML_WhitePaper.pdf)",
    )
    parser.add_argument(
        "--pages",
        type=int,
        default=2,
        help="Number of leading pages to print (default: 2)",
    )
    args = parser.parse_args()

    content = args.pdf_path.read_bytes()
    payload = json.loads(extract_ir_json(content, format="pdf"))

    pages = payload.get("pdf", {}).get("pages", [])
    selected = pages[: max(args.pages, 0)]

    if not selected:
        print("No extracted markdown pages found.")
        return

    for idx, page in enumerate(selected):
        number = page.get("number", idx + 1)
        markdown = page.get("markdown", "")
        print(f"## Page: {number}\n")
        print(markdown)
        if idx + 1 < len(selected):
            print("\n---\n")


if __name__ == "__main__":
    main()
