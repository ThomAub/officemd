# /// script
# requires-python = ">=3.12"
# dependencies = []
# ///

"""Bump all package versions, commit, and tag.

Usage:
    uv run bump.py patch          # 0.1.4 -> 0.1.5
    uv run bump.py minor          # 0.1.4 -> 0.2.0
    uv run bump.py major          # 0.1.4 -> 1.0.0
    uv run bump.py patch --dry-run
"""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent

# Every file that contains a version string to bump.
# Tuples of (relative_path, pattern_template, replacement_template).
# {old} and {new} are substituted at runtime.
VERSION_FILES: list[tuple[str, str, str]] = [
    # Workspace Cargo.toml
    ("Cargo.toml", 'version = "{old}"', 'version = "{new}"'),
    # Rust crates
    ("crates/officemd_core/Cargo.toml", 'version = "{old}"', 'version = "{new}"'),
    ("crates/officemd_markdown/Cargo.toml", 'version = "{old}"', 'version = "{new}"'),
    ("crates/officemd_docx/Cargo.toml", 'version = "{old}"', 'version = "{new}"'),
    ("crates/officemd_xlsx/Cargo.toml", 'version = "{old}"', 'version = "{new}"'),
    ("crates/officemd_csv/Cargo.toml", 'version = "{old}"', 'version = "{new}"'),
    ("crates/officemd_pptx/Cargo.toml", 'version = "{old}"', 'version = "{new}"'),
    ("crates/officemd_pdf/Cargo.toml", 'version = "{old}"', 'version = "{new}"'),
    ("crates/officemd_docling/Cargo.toml", 'version = "{old}"', 'version = "{new}"'),
    ("crates/officemd_cli/Cargo.toml", 'version = "{old}"', 'version = "{new}"'),
    ("crates/officemd_js/Cargo.toml", 'version = "{old}"', 'version = "{new}"'),
    ("crates/officemd_python/Cargo.toml", 'version = "{old}"', 'version = "{new}"'),
    # JS package.json (main + platform packages)
    ("crates/officemd_js/package.json", '"{old}"', '"{new}"'),
    ("crates/officemd_js/npm/darwin-arm64/package.json", '"{old}"', '"{new}"'),
    ("crates/officemd_js/npm/darwin-x64/package.json", '"{old}"', '"{new}"'),
    ("crates/officemd_js/npm/linux-arm64-gnu/package.json", '"{old}"', '"{new}"'),
    ("crates/officemd_js/npm/linux-arm64-musl/package.json", '"{old}"', '"{new}"'),
    ("crates/officemd_js/npm/linux-x64-gnu/package.json", '"{old}"', '"{new}"'),
    ("crates/officemd_js/npm/linux-x64-musl/package.json", '"{old}"', '"{new}"'),
    ("crates/officemd_js/npm/win32-arm64-msvc/package.json", '"{old}"', '"{new}"'),
    ("crates/officemd_js/npm/win32-x64-msvc/package.json", '"{old}"', '"{new}"'),
    # Python pyproject.toml
    ("crates/officemd_python/pyproject.toml", 'version = "{old}"', 'version = "{new}"'),
]


def get_current_version() -> str:
    """Read current version from workspace Cargo.toml."""
    cargo = ROOT / "Cargo.toml"
    match = re.search(r'version\s*=\s*"(\d+\.\d+\.\d+)"', cargo.read_text())
    if not match:
        print("error: could not find version in Cargo.toml", file=sys.stderr)
        sys.exit(1)
    return match.group(1)


def bump_version(current: str, part: str) -> str:
    major, minor, patch = (int(x) for x in current.split("."))
    if part == "major":
        return f"{major + 1}.0.0"
    if part == "minor":
        return f"{major}.{minor + 1}.0"
    return f"{major}.{minor}.{patch + 1}"


def replace_in_file(path: Path, old: str, new: str, dry_run: bool) -> int:
    """Replace all occurrences of old with new. Returns count of replacements."""
    if not path.exists():
        print(f"  skip (not found): {path.relative_to(ROOT)}")
        return 0
    text = path.read_text()
    count = text.count(old)
    if count == 0:
        print(f"  skip (no match):  {path.relative_to(ROOT)}")
        return 0
    if not dry_run:
        path.write_text(text.replace(old, new))
    print(f"  {path.relative_to(ROOT)}: {count} replacement(s)")
    return count


def run(cmd: list[str], check: bool = True) -> subprocess.CompletedProcess[str]:
    print(f"$ {' '.join(cmd)}")
    return subprocess.run(cmd, cwd=ROOT, check=check, capture_output=True, text=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="Bump all package versions, commit, and tag.")
    parser.add_argument("part", choices=["major", "minor", "patch"])
    parser.add_argument("--dry-run", action="store_true", help="Show what would change without writing")
    args = parser.parse_args()

    old = get_current_version()
    new = bump_version(old, args.part)
    tag = f"v{new}"

    print(f"{'[dry-run] ' if args.dry_run else ''}Bumping {old} -> {new}\n")

    total = 0
    for rel_path, pattern_tpl, replace_tpl in VERSION_FILES:
        path = ROOT / rel_path
        old_str = pattern_tpl.format(old=old, new=new)
        new_str = replace_tpl.format(old=old, new=new)
        total += replace_in_file(path, old_str, new_str, args.dry_run)

    if total == 0:
        print("\nerror: no replacements made - is the version already bumped?", file=sys.stderr)
        sys.exit(1)

    print(f"\n{total} total replacement(s) across {len(VERSION_FILES)} files")

    if args.dry_run:
        print(f"\n[dry-run] Would commit and tag {tag}")
        return

    # Update Cargo.lock
    print("\nUpdating Cargo.lock...")
    result = run(["cargo", "check", "--workspace"], check=False)
    if result.returncode != 0:
        print(f"error: cargo check failed:\n{result.stderr}", file=sys.stderr)
        sys.exit(1)

    # Commit
    print(f"\nCommitting...")
    run(["git", "add", "-A"])
    run(["git", "commit", "-m", f"Bump all packages to {tag}"])

    # Tag
    print(f"\nTagging {tag}...")
    run(["git", "tag", tag])

    print(f"\nDone. To push:\n  git push origin main && git push origin {tag}")


if __name__ == "__main__":
    main()
