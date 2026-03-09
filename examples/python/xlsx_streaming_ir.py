#!/usr/bin/env python3
import argparse
import json
from pathlib import Path

from officemd import extract_tables_ir_json


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract XLSX table IR JSON with streaming_rows enabled."
    )
    parser.add_argument("xlsx_path", type=Path, help="Path to .xlsx file")
    parser.add_argument(
        "--style-aware-values",
        action="store_true",
        help="Enable style-aware numeric/date display rendering.",
    )
    args = parser.parse_args()

    content = args.xlsx_path.read_bytes()
    payload = extract_tables_ir_json(
        content,
        style_aware_values=args.style_aware_values,
        streaming_rows=True,
    )
    print(json.dumps(json.loads(payload), indent=2))


if __name__ == "__main__":
    main()
