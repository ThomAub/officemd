# /// script
# requires-python = ">=3.12"
# dependencies = []
# ///
"""Generate a formatted benchmark report from hyperfine JSON results.

Reads all *.json files in the results directory (hyperfine output) and produces
a summary table with median latency and speedup ratios.

Usage:
    uv run benchmark/report.py --results-dir benchmark/results
    uv run benchmark/report.py --results-dir benchmark/results --json
"""

import argparse
import json
import sys
from pathlib import Path


def load_results(results_dir: Path) -> dict[str, dict[str, float]]:
    """Load hyperfine JSON results.

    Returns {filename: {tool_name: median_seconds}}.
    """
    data: dict[str, dict[str, float]] = {}

    for json_file in sorted(results_dir.glob("*.json")):
        with open(json_file) as f:
            try:
                raw = json.load(f)
            except json.JSONDecodeError:
                continue

        # Filename is e.g. "showcase.csv.json" -> stem is "showcase.csv"
        filename = json_file.stem
        tools: dict[str, float] = {}
        for result in raw.get("results", []):
            name = result.get("command", result.get("parameter", "unknown"))
            # hyperfine uses "command" field; --command-name sets it
            # but the JSON stores the display name in "command" when -n is used
            median = result.get("median", 0.0)
            tools[name] = median

        data[filename] = tools

    return data


def format_ms(seconds: float) -> str:
    """Format seconds as milliseconds string."""
    ms = seconds * 1000
    if ms < 0.01:
        return "<0.1ms"
    if ms < 1:
        return f"{ms:.2f}ms"
    if ms < 10:
        return f"{ms:.1f}ms"
    return f"{ms:.0f}ms"


def format_speedup(ratio: float) -> str:
    """Format speedup ratio."""
    if ratio < 10:
        return f"{ratio:.1f}x"
    return f"{ratio:.0f}x"


def discover_tools(data: dict[str, dict[str, float]]) -> list[str]:
    """Discover all tool names across results, ordered: officemd first, then alphabetical."""
    tools: set[str] = set()
    for file_tools in data.values():
        tools.update(file_tools.keys())

    ordered = []
    if "officemd" in tools:
        ordered.append("officemd")
        tools.discard("officemd")
    ordered.extend(sorted(tools))
    return ordered


def build_speed_table(
    data: dict[str, dict[str, float]], tools: list[str]
) -> list[list[str]]:
    """Build the speed comparison table rows.

    Returns list of rows, each row is [filename, tool1_ms, tool2_ms, ..., speedup].
    """
    rows = []
    for filename in sorted(data.keys()):
        file_tools = data[filename]
        row = [filename]

        officemd_median = file_tools.get("officemd")

        for tool in tools:
            median = file_tools.get(tool)
            if median is not None:
                row.append(format_ms(median))
            else:
                row.append("-")

        # Speedup: markitdown / officemd
        markitdown_median = file_tools.get("markitdown")
        if officemd_median and markitdown_median and officemd_median > 0.0001:
            ratio = markitdown_median / officemd_median
            row.append(format_speedup(ratio))
        else:
            row.append("-")

        rows.append(row)

    return rows


def render_table(headers: list[str], rows: list[list[str]], right_align: set[int] | None = None) -> str:
    """Render a markdown-compatible aligned table."""
    if right_align is None:
        right_align = set(range(1, len(headers)))

    widths = [len(h) for h in headers]
    for row in rows:
        for i, cell in enumerate(row):
            if i < len(widths):
                widths[i] = max(widths[i], len(cell))

    lines = []

    # Header
    header_cells = []
    for i, h in enumerate(headers):
        if i in right_align:
            header_cells.append(h.rjust(widths[i]))
        else:
            header_cells.append(h.ljust(widths[i]))
    lines.append("| " + " | ".join(header_cells) + " |")

    # Separator
    sep_cells = []
    for i in range(len(headers)):
        if i in right_align:
            sep_cells.append("-" * (widths[i] - 1) + ":")
        else:
            sep_cells.append("-" * widths[i])
    lines.append("| " + " | ".join(sep_cells) + " |")

    # Data rows
    for row in rows:
        cells = []
        for i in range(len(headers)):
            val = row[i] if i < len(row) else ""
            if i in right_align:
                cells.append(val.rjust(widths[i]))
            else:
                cells.append(val.ljust(widths[i]))
        lines.append("| " + " | ".join(cells) + " |")

    return "\n".join(lines)


def generate_report(data: dict[str, dict[str, float]]) -> str:
    """Generate the full benchmark report as markdown."""
    tools = discover_tools(data)

    sections = []
    sections.append("# officemd benchmark - text parsers and PDF to markdown")
    sections.append("")

    # Speed comparison
    headers = ["file"] + tools + ["speedup vs markitdown"]
    rows = build_speed_table(data, tools)

    sections.append("## Speed comparison - Median latency")
    sections.append("")
    sections.append(render_table(headers, rows))
    sections.append("")

    # Summary stats
    officemd_total = 0.0
    markitdown_total = 0.0
    count = 0
    for file_tools in data.values():
        o = file_tools.get("officemd")
        m = file_tools.get("markitdown")
        if o is not None and m is not None:
            officemd_total += o
            markitdown_total += m
            count += 1

    if count > 0 and officemd_total > 0:
        overall_speedup = markitdown_total / officemd_total
        sections.append(f"Overall: officemd is {format_speedup(overall_speedup)} faster "
                        f"than markitdown across {count} files.")

        # Per-tool aggregate stats
        for tool in ["markit", "liteparse"]:
            tool_officemd = 0.0
            tool_total = 0.0
            tool_count = 0
            for file_tools in data.values():
                o = file_tools.get("officemd")
                t = file_tools.get(tool)
                if o is not None and t is not None:
                    tool_officemd += o
                    tool_total += t
                    tool_count += 1

            if tool_count > 0 and tool_officemd > 0:
                speedup = tool_total / tool_officemd
                label = "PDF files" if tool == "liteparse" else "All files"
                sections.append(f"{label}: officemd is {format_speedup(speedup)} faster "
                                f"than {tool} across {tool_count} files.")

    sections.append("")
    return "\n".join(sections)


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate benchmark report from hyperfine results.")
    parser.add_argument(
        "--results-dir",
        type=Path,
        default=Path(__file__).resolve().parent / "results",
        help="Directory containing hyperfine JSON files",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Write report to file (default: stdout only)",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Output raw data as JSON instead of markdown report",
    )
    args = parser.parse_args()

    if not args.results_dir.is_dir():
        print(f"Results directory not found: {args.results_dir}", file=sys.stderr)
        return 1

    data = load_results(args.results_dir)
    if not data:
        print("No hyperfine JSON results found.", file=sys.stderr)
        return 1

    if args.json:
        output = json.dumps(data, indent=2)
    else:
        output = generate_report(data)

    print(output)

    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(output + "\n")
        print(f"\nReport written to {args.output}", file=sys.stderr)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
