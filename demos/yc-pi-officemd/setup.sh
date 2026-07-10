#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
REPO_ROOT=$(cd -- "$SCRIPT_DIR/../.." && pwd)
FIXTURES_DIR="$SCRIPT_DIR/fixtures"
INPUTS_DIR="$FIXTURES_DIR/inputs"

mkdir -p "$INPUTS_DIR"

download() {
  local url=$1
  local destination=$2
  local expected_sha256=$3
  local temporary="${destination}.download"

  if [[ -f "$destination" ]] && [[ $(shasum -a 256 "$destination" | awk '{print $1}') == "$expected_sha256" ]]; then
    printf 'ready  %s\n' "$(basename "$destination")"
    return
  fi

  rm -f "$temporary"
  curl --fail --location --retry 3 --silent --show-error --output "$temporary" "$url"
  local actual_sha256
  actual_sha256=$(shasum -a 256 "$temporary" | awk '{print $1}')
  if [[ "$actual_sha256" != "$expected_sha256" ]]; then
    rm -f "$temporary"
    printf 'checksum mismatch for %s: expected %s, got %s\n' "$destination" "$expected_sha256" "$actual_sha256" >&2
    exit 1
  fi
  mv "$temporary" "$destination"
  printf 'fetched %s\n' "$(basename "$destination")"
}

BASE_URL="https://huggingface.co/datasets/openai/gdpval/resolve/main"

download \
  "$BASE_URL/reference_files/0c9d7139ad82b8101a10705716fde830/Pick%20Tickets%20062525.xlsx" \
  "$INPUTS_DIR/Pick Tickets 062525.xlsx" \
  "0c9d7139ad82b8101a10705716fde8300b80400b6231022442ae37dada780f8d"

download \
  "$BASE_URL/reference_files/aa5b2c0f19996b0927ee429972fcfb93/Blank%20Daily%20Shipment%20Manifest.xlsx" \
  "$INPUTS_DIR/Blank Daily Shipment Manifest.xlsx" \
  "aa5b2c0f19996b0927ee429972fcfb93b6d7beece74dc4ae79f468148d0de5a3"

download \
  "$BASE_URL/reference_files/63edd16ae28e50b012347ea841b03c64/Shipping%20parameters.xlsx" \
  "$INPUTS_DIR/Shipping parameters.xlsx" \
  "63edd16ae28e50b012347ea841b03c64efd5d3f76e3564f2630068e9ba0e41bd"

printf 'building OfficeMD CLI\n'
cargo build --quiet --manifest-path "$REPO_ROOT/Cargo.toml" -p officemd_cli
printf 'setup complete: %s\n' "$FIXTURES_DIR"
