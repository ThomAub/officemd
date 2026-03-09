#!/usr/bin/env python3
import argparse
import json
from pathlib import Path

from officemd import inspect_pdf_json

DEFAULT_SCANNED = Path("examples/data/ocr_graph_scanned.pdf")
DEFAULT_OCRED = Path("examples/data/ocr_graph_ocred.pdf")


def inspect(path: Path) -> dict:
    return json.loads(inspect_pdf_json(path.read_bytes()))


def ocr_ratio(diagnostics: dict) -> float:
    page_count = int(diagnostics.get("page_count") or 0)
    if page_count <= 0:
        return 0.0
    pages_needing_ocr = diagnostics.get("pages_needing_ocr") or []
    return len(pages_needing_ocr) / page_count


def print_compact(label: str, diagnostics: dict) -> None:
    pages_needing_ocr = diagnostics.get("pages_needing_ocr") or []
    confidence = float(diagnostics.get("confidence") or 0.0)
    print(
        f"{label}: "
        f"classification={diagnostics.get('classification')} "
        f"confidence={confidence:.4f} "
        f"page_count={diagnostics.get('page_count')} "
        f"pages_needing_ocr={pages_needing_ocr} "
        f"ocr_ratio={ocr_ratio(diagnostics):.2f} "
        f"has_encoding_issues={diagnostics.get('has_encoding_issues')}"
    )


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Compare OCR-readiness diagnostics for two PDF fixtures."
    )
    parser.add_argument(
        "--scanned",
        type=Path,
        default=DEFAULT_SCANNED,
        help=f"Path to scanned-like PDF (default: {DEFAULT_SCANNED})",
    )
    parser.add_argument(
        "--ocred",
        type=Path,
        default=DEFAULT_OCRED,
        help=f"Path to OCRed/text PDF (default: {DEFAULT_OCRED})",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Print full JSON diagnostics for both files",
    )
    args = parser.parse_args()

    scanned = inspect(args.scanned)
    ocred = inspect(args.ocred)

    if args.json:
        print(
            json.dumps(
                {
                    "scanned": {"path": str(args.scanned), "diagnostics": scanned},
                    "ocred": {"path": str(args.ocred), "diagnostics": ocred},
                },
                indent=2,
            )
        )
        return

    print_compact("scanned", scanned)
    print_compact("ocred", ocred)


if __name__ == "__main__":
    main()
