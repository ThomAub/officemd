#!/usr/bin/env python3
from __future__ import annotations

import argparse
import html
import json
import os
from datetime import datetime
from pathlib import Path
from urllib.parse import quote


def read_json(path: Path) -> dict:
    if not path.is_file():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def relative_url(path: Path, root: Path) -> str:
    return quote(os.path.relpath(path, root).replace(os.sep, "/"))


def process_metrics(root: Path, run: dict) -> dict[str, str | int]:
    tool_calls = 0
    officemd_calls = 0
    install_calls = 0
    events = root / "events.jsonl"
    if events.is_file():
        for line in events.read_text(encoding="utf-8").splitlines():
            try:
                event = json.loads(line)
            except json.JSONDecodeError:
                continue
            if event.get("type") != "tool_execution_start":
                continue
            tool_calls += 1
            command = str(event.get("args", {}).get("command", ""))
            if event.get("toolName") == "bash" and "officemd" in command:
                officemd_calls += 1
            if event.get("toolName") == "bash" and (
                "pip install" in command or "npm install" in command
            ):
                install_calls += 1

    elapsed = "unknown"
    try:
        started = datetime.fromisoformat(str(run["started"]).replace("Z", "+00:00"))
        ended = datetime.fromisoformat(str(run["ended"]).replace("Z", "+00:00"))
        elapsed_seconds = int((ended - started).total_seconds())
        elapsed = f"{elapsed_seconds // 60}m {elapsed_seconds % 60}s"
    except (KeyError, TypeError, ValueError):
        pass

    output = root / "work" / "Daily Shipment Manifest 062525.xlsx"
    return {
        "elapsed": elapsed,
        "tool_calls": tool_calls,
        "officemd_calls": officemd_calls,
        "install_calls": install_calls,
        "provenance": "yes" if output.with_name(output.name + ".officemd.json").is_file() else "no",
    }


def render_variant(run_dir: Path, variant: str) -> str:
    root = run_dir / variant
    score = read_json(root / "score.json")
    run = read_json(root / "run.json")
    metrics = process_metrics(root, run)
    criteria = score.get("criteria", [])
    rows = "".join(
        "<tr>"
        f"<td>{html.escape(item['name'])}</td>"
        f"<td class='num'>{item['score']:g}/{item['maximum']:g}</td>"
        f"<td>{html.escape(item['detail'])}</td>"
        "</tr>"
        for item in criteria
    )
    images = sorted((root / "evidence" / "render").glob("*.png"))
    preview = (
        f"<img src='{relative_url(images[0], run_dir)}' alt='{variant} workbook render'>"
        if images
        else "<p class='missing'>No rendered evidence</p>"
    )
    return f"""
    <section class="variant">
      <header>
        <div>
          <p class="label">{html.escape(variant)}</p>
          <h2>{score.get('score', 0):g}<span>/100</span></h2>
        </div>
        <dl>
          <dt>Model</dt><dd>{html.escape(str(run.get('provider', 'unknown')))}/{html.escape(str(run.get('model', 'unknown')))}</dd>
          <dt>Exit</dt><dd>{html.escape(str(run.get('exit_code', 'unknown')))}</dd>
        </dl>
      </header>
      <div class="control" aria-label="Artifact control evidence">
        <div><span>Typed provenance</span><strong>{metrics['provenance']}</strong></div>
        <div><span>Package installs</span><strong>{metrics['install_calls']}</strong></div>
        <div><span>OfficeMD calls</span><strong>{metrics['officemd_calls']}</strong></div>
        <div><span>Elapsed</span><strong>{metrics['elapsed']}</strong></div>
      </div>
      <div class="preview">{preview}</div>
      <table>
        <thead><tr><th>Criterion</th><th>Score</th><th>Evidence</th></tr></thead>
        <tbody>{rows}</tbody>
      </table>
    </section>
    """


def main() -> None:
    parser = argparse.ArgumentParser(description="Build the OfficeMD YC demo comparison report")
    parser.add_argument("--run-dir", type=Path, required=True)
    args = parser.parse_args()
    run_dir = args.run_dir.resolve()
    variants = [name for name in ("baseline", "skill") if (run_dir / name / "score.json").is_file()]
    body = "".join(render_variant(run_dir, variant) for variant in variants)
    document = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Pi + OfficeMD GDPval comparison</title>
  <style>
    :root {{ color-scheme: light; font-family: Inter, ui-sans-serif, system-ui, sans-serif; color: #171717; background: #f5f5f3; }}
    * {{ box-sizing: border-box; }}
    body {{ margin: 0; }}
    main {{ width: min(1500px, 100%); margin: 0 auto; padding: 32px; }}
    .title {{ display: flex; justify-content: space-between; align-items: end; gap: 24px; margin-bottom: 24px; border-bottom: 2px solid #171717; padding-bottom: 16px; }}
    h1 {{ margin: 0; font-family: Georgia, serif; font-size: 34px; letter-spacing: 0; }}
    .title p {{ margin: 0; color: #666; text-align: right; }}
    .comparison {{ display: grid; grid-template-columns: repeat({max(len(variants), 1)}, minmax(0, 1fr)); border: 1px solid #c9c9c5; background: white; }}
    .variant {{ min-width: 0; padding: 20px; }}
    .variant + .variant {{ border-left: 1px solid #c9c9c5; }}
    .variant header {{ display: flex; justify-content: space-between; gap: 20px; align-items: start; min-height: 92px; }}
    .label {{ margin: 0; text-transform: uppercase; font-size: 12px; font-weight: 700; color: #007f5f; }}
    h2 {{ margin: 2px 0 0; font-size: 48px; letter-spacing: 0; }}
    h2 span {{ font-size: 18px; color: #777; }}
    dl {{ margin: 0; display: grid; grid-template-columns: auto auto; gap: 4px 12px; font-size: 12px; }}
    dt {{ color: #777; }} dd {{ margin: 0; font-weight: 600; }}
    .control {{ display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); border-top: 1px solid #d7d7d3; border-bottom: 1px solid #d7d7d3; margin: 2px 0 14px; }}
    .control div {{ min-width: 0; padding: 9px 8px; }}
    .control div + div {{ border-left: 1px solid #d7d7d3; }}
    .control span, .control strong {{ display: block; overflow-wrap: anywhere; }}
    .control span {{ color: #777; font-size: 10px; margin-bottom: 3px; }}
    .control strong {{ font-size: 13px; }}
    .preview {{ aspect-ratio: 4 / 3; background: #ececea; border: 1px solid #d7d7d3; display: grid; place-items: center; overflow: hidden; margin: 14px 0 18px; }}
    .preview img {{ width: 100%; height: 100%; object-fit: contain; }}
    .missing {{ color: #777; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
    th {{ text-align: left; border-bottom: 2px solid #333; padding: 7px 6px; }}
    td {{ border-bottom: 1px solid #dededb; padding: 7px 6px; vertical-align: top; }}
    td:last-child {{ color: #666; }}
    .num {{ white-space: nowrap; font-variant-numeric: tabular-nums; font-weight: 700; }}
    footer {{ color: #666; font-size: 12px; margin-top: 14px; }}
    @media (max-width: 900px) {{ main {{ padding: 16px; }} .title {{ display: block; }} .title p {{ text-align: left; margin-top: 8px; }} .comparison {{ grid-template-columns: 1fr; }} .variant + .variant {{ border-left: 0; border-top: 1px solid #c9c9c5; }} }}
  </style>
</head>
<body>
  <main>
    <div class="title">
      <h1>Same Pi. Same task. Better artifact control.</h1>
      <p>GDPval task 76418a2c · objective XLSX rubric · native render evidence</p>
    </div>
    <div class="comparison">{body}</div>
    <footer>Scores come from workbook cells, formulas, source hashes, OfficeMD verification, and rendered PNG evidence.</footer>
  </main>
</body>
</html>
"""
    output = run_dir / "report.html"
    output.write_text(document, encoding="utf-8")
    print(output)


if __name__ == "__main__":
    main()
