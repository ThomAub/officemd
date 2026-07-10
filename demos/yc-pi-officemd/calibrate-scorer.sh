#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
REPO_ROOT=$(cd -- "$SCRIPT_DIR/../.." && pwd)
OFFICEMD="$REPO_ROOT/target/debug/officemd"
TEMP_ROOT=$(mktemp -d "/tmp/officemd-yc-calibration.XXXXXX")
trap 'rm -rf "$TEMP_ROOT"' EXIT

"$SCRIPT_DIR/setup.sh"
mkdir -p "$TEMP_ROOT/blank/work" "$TEMP_ROOT/reference/work"
cp "$SCRIPT_DIR/fixtures/inputs"/*.xlsx "$TEMP_ROOT/blank/work/"
cp "$SCRIPT_DIR/fixtures/inputs"/*.xlsx "$TEMP_ROOT/reference/work/"

REFERENCE_URL="https://huggingface.co/datasets/openai/gdpval/resolve/main/deliverable_files/51fd7afc9aecdf5650a1d6ea3498f2fd/Daily%20Shipment%20Manifest%20062525.xlsx"
REFERENCE_OUTPUT="$TEMP_ROOT/reference/work/Daily Shipment Manifest 062525.xlsx"
curl --fail --location --retry 3 --silent --show-error --output "$REFERENCE_OUTPUT" "$REFERENCE_URL"

expected_sha256="179aef0a5befc3636d7ffd0886a6eebddad7b18368640b74bf537ad81ecb03b2"
actual_sha256=$(shasum -a 256 "$REFERENCE_OUTPUT" | awk '{print $1}')
if [[ "$actual_sha256" != "$expected_sha256" ]]; then
  printf 'reference checksum mismatch: expected %s, got %s\n' "$expected_sha256" "$actual_sha256" >&2
  exit 1
fi

for variant in blank reference; do
  uv run "$SCRIPT_DIR/score.py" \
    --officemd "$OFFICEMD" \
    --work-dir "$TEMP_ROOT/$variant/work" \
    --fixtures-dir "$SCRIPT_DIR/fixtures/inputs" \
    --evidence-dir "$TEMP_ROOT/$variant/evidence" \
    --output "$TEMP_ROOT/$variant/score.json"
done

blank_score=$(jq -r '.score' "$TEMP_ROOT/blank/score.json")
reference_score=$(jq -r '.score' "$TEMP_ROOT/reference/score.json")
printf 'blank=%s/100 reference=%s/100\n' "$blank_score" "$reference_score"

[[ "$blank_score" == "5.0" ]]
[[ "$reference_score" == "100.0" ]]
