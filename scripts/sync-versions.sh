#!/usr/bin/env bash
# Sync the workspace version from Cargo.toml to pyproject.toml and package.json.
# Called by release-plz workflow after the release PR is created/updated.

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "$0")/.." && pwd)"

# Extract workspace version from root Cargo.toml
VERSION=$(grep -m1 '^version = ' "$REPO_ROOT/Cargo.toml" | sed 's/version = "\(.*\)"/\1/')

if [ -z "$VERSION" ]; then
    echo "ERROR: Could not extract version from Cargo.toml"
    exit 1
fi

echo "Syncing version: $VERSION"

# Update pyproject.toml
PYPROJECT="$REPO_ROOT/crates/officemd_python/pyproject.toml"
if [ -f "$PYPROJECT" ]; then
    sed -i.bak "s/^version = \".*\"/version = \"$VERSION\"/" "$PYPROJECT"
    rm -f "$PYPROJECT.bak"
    echo "Updated $PYPROJECT"
fi

# Update package.json (top-level version + optionalDependencies versions)
PACKAGEJSON="$REPO_ROOT/crates/officemd_js/package.json"
if [ -f "$PACKAGEJSON" ]; then
    jq --arg v "$VERSION" '
      .version = $v |
      if .optionalDependencies then .optionalDependencies |= map_values($v) else . end
    ' "$PACKAGEJSON" > "$PACKAGEJSON.tmp" && mv "$PACKAGEJSON.tmp" "$PACKAGEJSON"
    echo "Updated $PACKAGEJSON"
fi

echo "Version sync complete: $VERSION"
