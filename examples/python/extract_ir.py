#!/usr/bin/env python3
import argparse
import json
from pathlib import Path

from officemd import detect_format, extract_ir_json


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract IR JSON from a DOCX/XLSX/CSV/PPTX file."
    )
    parser.add_argument(
        "input_path", type=Path, help="Path to .docx/.xlsx/.csv/.pptx file"
    )
    parser.add_argument(
        "--format",
        choices=["docx", "xlsx", "csv", "pptx"],
        help="Optional explicit format override",
    )
    args = parser.parse_args()

    content = args.input_path.read_bytes()
    if args.format is None and args.input_path.suffix.lower() != ".csv":
        print(f"detected_format={detect_format(content)}")

    payload = extract_ir_json(content, format=args.format)
    print(json.dumps(json.loads(payload), indent=2))


if __name__ == "__main__":
    main()
