# /// script
# requires-python = ">=3.12"
# dependencies = ["docling-core>=1.9.0", "officemd"]
# ///
"""Compare our Rust docling conversion against docling-core's own output.

Converts showcase/bench files with both our officemd.docling_from_bytes binding
and docling-core, then diffs outputs and saves snapshots.

Usage:
    uv run examples/python/compare_docling.py
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from officemd import docling_from_bytes


ROOT = Path(__file__).resolve().parents[2]
DATA_DIR = ROOT / "examples" / "data"
BENCH_DIR = ROOT / "examples" / "bench-data"
SNAPSHOT_DIR = ROOT / "tests" / "snapshots"


def canonical_json(payload: str) -> dict:
    """Parse and return a normalized dict for comparison."""
    return json.loads(payload)


def collect_files() -> list[Path]:
    """Gather all docx/xlsx/pptx files from data and bench-data dirs."""
    files = []
    for d in [DATA_DIR, BENCH_DIR]:
        if not d.exists():
            continue
        for ext in ("*.docx", "*.xlsx", "*.pptx"):
            files.extend(sorted(d.glob(ext)))
    return files


def format_for_path(path: Path) -> str:
    return path.suffix.lstrip(".")


def convert_with_rust(path: Path) -> dict:
    content = path.read_bytes()
    fmt = format_for_path(path)
    payload = docling_from_bytes(content, format=fmt)
    return canonical_json(payload)


def structural_diff(ours: dict, theirs: dict, path: str = "") -> list[str]:
    """Recursively compare two dicts, returning a list of diff descriptions."""
    diffs = []

    if type(ours) is not type(theirs):
        diffs.append(
            f"{path}: type mismatch: {type(ours).__name__} vs {type(theirs).__name__}"
        )
        return diffs

    if isinstance(ours, dict):
        all_keys = sorted(set(ours) | set(theirs))
        for key in all_keys:
            child_path = f"{path}.{key}" if path else key
            if key not in ours:
                diffs.append(f"{child_path}: missing in ours")
            elif key not in theirs:
                diffs.append(f"{child_path}: missing in theirs")
            else:
                diffs.extend(structural_diff(ours[key], theirs[key], child_path))
    elif isinstance(ours, list):
        if len(ours) != len(theirs):
            diffs.append(f"{path}: list length {len(ours)} vs {len(theirs)}")
        for i, (a, b) in enumerate(zip(ours, theirs)):
            diffs.extend(structural_diff(a, b, f"{path}[{i}]"))
    elif ours != theirs:
        ours_str = str(ours)[:80]
        theirs_str = str(theirs)[:80]
        diffs.append(f"{path}: {ours_str!r} vs {theirs_str!r}")

    return diffs


def main() -> None:
    files = collect_files()
    if not files:
        print("No test files found")
        sys.exit(1)

    print(f"Found {len(files)} files to compare\n")

    # Try to import docling for comparison
    docling_available = False
    try:
        from docling.document_converter import DocumentConverter  # ty: ignore[unresolved-import]

        docling_available = True
        converter = DocumentConverter()
    except ImportError:
        print(
            "docling not installed - skipping comparison, only saving Rust snapshots\n"
        )

    summary: list[tuple[str, int, str]] = []

    for path in files:
        name = path.name
        print(f"--- {name} ---")

        # Our Rust conversion
        try:
            ours = convert_with_rust(path)
            rust_ok = True
        except Exception as e:
            print(f"  Rust conversion failed: {e}")
            summary.append((name, -1, f"Rust error: {e}"))
            rust_ok = False
            continue

        # Save snapshot
        snapshot_name = f"{path.stem}_{format_for_path(path)}_docling.json"
        snapshot_path = SNAPSHOT_DIR / snapshot_name
        snapshot_path.write_text(
            json.dumps(ours, indent=2, sort_keys=True, ensure_ascii=False) + "\n"
        )
        print(f"  Snapshot saved: {snapshot_path.name}")

        if docling_available and rust_ok:
            try:
                result = converter.convert(str(path))
                theirs_json = result.document.export_to_dict()
                theirs = json.loads(json.dumps(theirs_json))  # normalize
                diffs = structural_diff(ours, theirs)
                n = len(diffs)
                summary.append((name, n, ""))
                if n == 0:
                    print("  Match: identical")
                else:
                    print(f"  Diffs: {n}")
                    for d in diffs[:10]:
                        print(f"    {d}")
                    if n > 10:
                        print(f"    ... and {n - 10} more")
            except Exception as e:
                print(f"  docling conversion failed: {e}")
                summary.append((name, -1, f"docling error: {e}"))
        elif rust_ok:
            summary.append((name, -2, "docling not available"))

    # Summary table
    print("\n=== Summary ===")
    print(f"{'File':<40} {'Diffs':>8} {'Notes'}")
    print("-" * 70)
    for name, n, notes in summary:
        if n == -1:
            status = "ERROR"
        elif n == -2:
            status = "N/A"
        elif n == 0:
            status = "MATCH"
        else:
            status = str(n)
        print(f"{name:<40} {status:>8} {notes}")


if __name__ == "__main__":
    main()
