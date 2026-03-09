#!/usr/bin/env python3
import argparse
import json
from pathlib import Path

from officemd import inspect_pdf_fonts_json


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Inspect detected PDF fonts and print a compact summary."
    )
    parser.add_argument(
        "pdf_path",
        type=Path,
        nargs="?",
        default=Path("examples/data/OpenXML_WhitePaper.pdf"),
        help="Path to .pdf file (default: examples/data/OpenXML_WhitePaper.pdf)",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=10,
        help="Maximum number of fonts to print in summary mode (default: 10)",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Print full inspection JSON output",
    )
    args = parser.parse_args()

    payload = json.loads(inspect_pdf_fonts_json(args.pdf_path.read_bytes()))

    if args.json:
        print(json.dumps(payload, indent=2))
        return

    diagnostics = payload.get("diagnostics", {})
    print(
        "classification=",
        diagnostics.get("classification"),
        "page_count=",
        diagnostics.get("page_count"),
        "fonts=",
        len(payload.get("fonts", [])),
    )

    fonts = payload.get("fonts", [])[: max(args.limit, 0)]
    for font in fonts:
        print(
            f"- {font.get('font_name')} "
            f"(text_items={font.get('text_item_count')}, pages={font.get('pages')})"
        )


if __name__ == "__main__":
    main()
