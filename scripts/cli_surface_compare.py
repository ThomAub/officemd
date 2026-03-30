#!/usr/bin/env python3
"""Compare public CLI surface across officemd implementations.

Examples:
  python scripts/cli_surface_compare.py
  python scripts/cli_surface_compare.py --impl rust js --skip-unavailable
  python scripts/cli_surface_compare.py --case "--help" --case "markdown --help"
  python scripts/cli_surface_compare.py \
      --js-cmd "node crates/officemd_js/cli.js" \
      --python-cmd "uv run --directory crates/officemd_python python -m officemd._cli"
"""

from __future__ import annotations

import argparse
import difflib
import os
import re
import shlex
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parent.parent
ANSI_RE = re.compile(r"\x1b\[[0-9;]*m")

DEFAULT_CASES = [
    "--help",
    "markdown --help",
    "render --help",
    "diff --help",
    "convert --help",
    "stream --help",
    "inspect --help",
    "create --help",
]


@dataclass(frozen=True)
class InvocationResult:
    implementation: str
    case: str
    argv: tuple[str, ...]
    exit_code: int
    stdout: str
    stderr: str
    error: str | None = None

    @property
    def ok(self) -> bool:
        return self.error is None and self.exit_code == 0

    def combined_output(self) -> str:
        return (
            f"exit_code: {self.exit_code}\n"
            f"----- stdout -----\n{self.stdout}"
            f"\n----- stderr -----\n{self.stderr}"
        )


@dataclass(frozen=True)
class CliSpec:
    name: str
    argv: tuple[str, ...]
    env: tuple[tuple[str, str], ...] = ()

    def command_for_case(self, case: str) -> tuple[str, ...]:
        return self.argv + tuple(shlex.split(case))


def default_specs() -> dict[str, CliSpec]:
    python_path = str(REPO_ROOT / "crates" / "officemd_python" / "python")
    return {
        "rust": CliSpec("rust", ("cargo", "run", "-q", "-p", "officemd_cli", "--")),
        "js": CliSpec("js", ("node", "crates/officemd_js/cli.js")),
        "python": CliSpec(
            "python",
            (sys.executable, "-m", "officemd._cli"),
            env=(("PYTHONPATH", python_path),),
        ),
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Compare officemd CLI help/error surface across implementations.",
    )
    parser.add_argument(
        "--impl",
        nargs="+",
        default=["rust", "js", "python"],
        choices=["rust", "js", "python"],
        help="Implementations to compare",
    )
    parser.add_argument(
        "--baseline",
        default="rust",
        choices=["rust", "js", "python"],
        help="Implementation used as the canonical output",
    )
    parser.add_argument(
        "--case",
        action="append",
        default=[],
        help="Invocation suffix to compare (repeatable). Defaults to standard help cases.",
    )
    parser.add_argument(
        "--rust-cmd",
        help="Override Rust launcher, e.g. 'cargo run -q -p officemd_cli --'",
    )
    parser.add_argument(
        "--js-cmd",
        help="Override JS launcher, e.g. 'node crates/officemd_js/cli.js'",
    )
    parser.add_argument(
        "--python-cmd",
        help="Override Python launcher, e.g. 'uv run --directory crates/officemd_python python -m officemd._cli'",
    )
    parser.add_argument(
        "--skip-unavailable",
        action="store_true",
        help="Skip implementations that fail to launch instead of failing immediately.",
    )
    return parser.parse_args()


def build_specs(args: argparse.Namespace) -> dict[str, CliSpec]:
    specs = default_specs()
    overrides = {
        "rust": args.rust_cmd,
        "js": args.js_cmd,
        "python": args.python_cmd,
    }
    for name, override in overrides.items():
        if override:
            specs[name] = CliSpec(name, tuple(shlex.split(override)))
    return specs


def normalize_output(text: str) -> str:
    text = ANSI_RE.sub("", text)
    text = text.replace("\r\n", "\n")
    text = text.replace("office-md", "officemd")
    text = text.replace("Usage: python -m officemd._cli", "Usage: officemd")
    text = text.replace("Usage: node crates/officemd_js/cli.js", "Usage: officemd")
    return text.strip() + "\n"


def run_case(spec: CliSpec, case: str) -> InvocationResult:
    argv = spec.command_for_case(case)
    env = os.environ.copy()
    for key, value in spec.env:
        existing = env.get(key)
        if key == "PYTHONPATH" and existing:
            env[key] = f"{value}{os.pathsep}{existing}"
        else:
            env[key] = value
    try:
        completed = subprocess.run(
            argv,
            cwd=REPO_ROOT,
            env=env,
            capture_output=True,
            text=True,
            check=False,
        )
    except OSError as exc:
        return InvocationResult(spec.name, case, argv, -1, "", "", error=str(exc))

    return InvocationResult(
        implementation=spec.name,
        case=case,
        argv=argv,
        exit_code=completed.returncode,
        stdout=normalize_output(completed.stdout),
        stderr=normalize_output(completed.stderr),
    )


def compare_case(case: str, baseline: InvocationResult, other: InvocationResult) -> str | None:
    baseline_text = baseline.combined_output().splitlines(keepends=True)
    other_text = other.combined_output().splitlines(keepends=True)
    if baseline_text == other_text:
        return None
    diff = difflib.unified_diff(
        baseline_text,
        other_text,
        fromfile=f"{baseline.implementation}:{case}",
        tofile=f"{other.implementation}:{case}",
    )
    return "".join(diff)


def print_result(result: InvocationResult) -> None:
    print(f"## {result.implementation}: {result.case}")
    print(f"$ {' '.join(shlex.quote(part) for part in result.argv)}")
    if result.error is not None:
        print(f"launch_error: {result.error}")
        return
    print(result.combined_output())


def main() -> int:
    args = parse_args()
    specs = build_specs(args)
    selected = args.impl
    cases = args.case or DEFAULT_CASES

    if args.baseline not in selected:
        print(f"error: baseline {args.baseline!r} must be included in --impl", file=sys.stderr)
        return 2

    results: dict[tuple[str, str], InvocationResult] = {}
    skipped: list[str] = []

    for impl in selected:
        spec = specs[impl]
        launch_probe = run_case(spec, cases[0])
        if launch_probe.error is not None or (launch_probe.exit_code != 0 and args.skip_unavailable):
            reason = launch_probe.error or launch_probe.stderr.strip() or f"exit {launch_probe.exit_code}"
            if args.skip_unavailable and impl != args.baseline:
                skipped.append(f"{impl}: {reason}")
                continue
            results[(impl, cases[0])] = launch_probe
            print_result(launch_probe)
            print()
            print("error: required implementation is unavailable", file=sys.stderr)
            return 1

        results[(impl, cases[0])] = launch_probe
        for case in cases[1:]:
            results[(impl, case)] = run_case(spec, case)

    if skipped:
        print("Skipped implementations:")
        for item in skipped:
            print(f"  - {item}")
        print()

    mismatches = 0
    failures = 0
    for case in cases:
        baseline = results[(args.baseline, case)]
        if not baseline.ok:
            print_result(baseline)
            print()
            print(f"error: baseline failed for case {case!r}", file=sys.stderr)
            return 1

        for impl in selected:
            if (impl, case) not in results or impl == args.baseline:
                continue
            result = results[(impl, case)]
            if not result.ok:
                failures += 1
                print_result(result)
                print()
                continue
            diff = compare_case(case, baseline, result)
            if diff is not None:
                mismatches += 1
                print(f"## mismatch: {case} ({args.baseline} vs {impl})")
                print(diff)

    if failures or mismatches:
        print(
            f"Surface comparison found {mismatches} mismatch(es) and {failures} failing invocation(s).",
            file=sys.stderr,
        )
        return 1

    print(
        f"Surface comparison passed for {', '.join(selected)} across {len(cases)} case(s)."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
