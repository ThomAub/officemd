#!/usr/bin/env bash
# Release OfficeMD to git and crates.io.
# PyPI and npm are published via CI (see .github/workflows/release.yml).
#
# Prerequisites:
#   cargo login
#
# Usage:
#   bash create_release.sh                      # full release
#   bash create_release.sh --dry-run            # print commands without executing
#   bash create_release.sh --start-from crates  # skip to crates.io publish

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "$0")" && pwd)"
VERSION="0.1.0"
DRY_RUN=false
START_FROM="preflight"  # preflight -> git -> crates -> verify

for arg in "$@"; do
    case "$arg" in
        --dry-run)    DRY_RUN=true ;;
        --start-from) ;;  # value handled below
        preflight|git|crates)
            START_FROM="$arg" ;;
        *)
            if [[ "${prev_arg:-}" == "--start-from" ]]; then
                START_FROM="$arg"
            else
                echo "Unknown argument: $arg" >&2
                exit 1
            fi
            ;;
    esac
    prev_arg="$arg"
done

if $DRY_RUN; then
    echo "=== DRY RUN MODE ==="
    echo ""
fi

if [[ "$START_FROM" != "preflight" ]]; then
    echo "=== Starting from: $START_FROM ==="
    echo ""
fi

# Steps in order; each step runs if START_FROM <= that step
STEPS=(preflight git crates verify)
step_reached() {
    local target="$1"
    for s in "${STEPS[@]}"; do
        if [[ "$s" == "$START_FROM" ]]; then return 0; fi
        if [[ "$s" == "$target" ]]; then return 1; fi
    done
    return 0
}

RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

log()  { echo -e "${GREEN}[+]${NC} $*"; }
warn() { echo -e "${YELLOW}[!]${NC} $*"; }
fail() { echo -e "${RED}[x]${NC} $*"; exit 1; }

run() {
    if $DRY_RUN; then
        echo "  (dry-run) $*"
    else
        "$@"
    fi
}

# ---------------------------------------------------------------------------
# Pre-flight checks
# ---------------------------------------------------------------------------

if step_reached preflight; then
    log "Pre-flight checks..."

    command -v cargo >/dev/null || fail "cargo not found"
    command -v git >/dev/null   || fail "git not found"
    command -v curl >/dev/null  || fail "curl not found"

    cd "$REPO_ROOT"

    CURRENT_TAG="$(git describe --tags --exact-match HEAD 2>/dev/null || true)"
    if [[ "$CURRENT_TAG" != "v$VERSION" ]]; then
        fail "HEAD is not tagged v$VERSION (got: '${CURRENT_TAG:-none}')"
    fi

    log "Running cargo nextest..."
    if ! $DRY_RUN; then
        cargo nextest run --workspace --cargo-quiet --failure-output final --status-level fail --fail-fast || fail "Tests failed"
    fi

    log "Running cargo clippy..."
    if ! $DRY_RUN; then
        cargo clippy --workspace --all-targets --exclude officemd_pdf --quiet -- -D warnings || fail "Clippy failed"
    fi

    log "Pre-flight checks passed."
    echo ""
fi

# ---------------------------------------------------------------------------
# Step 1: Push to git
# ---------------------------------------------------------------------------

if step_reached git; then
    log "Step 1: Push to origin..."
    run git push --force origin main --tags
    log "Git push done."
    echo ""
fi

# ---------------------------------------------------------------------------
# Step 2: Publish to crates.io in dependency order
# ---------------------------------------------------------------------------

# Publish order respects the dependency graph:
#   Level 1: officemd_core (no internal deps)
#   Level 2: officemd_markdown, officemd_docling (depend on core)
#   Level 3: officemd_docx, officemd_xlsx, officemd_csv, officemd_pptx, officemd_pdf
#   Level 4: officemd_cli (depends on everything)

CRATES_LEVEL_1=("officemd_core")
CRATES_LEVEL_2=("officemd_markdown" "officemd_docling")
CRATES_LEVEL_3=("officemd_docx" "officemd_xlsx" "officemd_csv" "officemd_pptx" "officemd_pdf")
CRATES_LEVEL_4=("officemd_cli")

wait_for_crate() {
    local crate="$1"
    local version="$2"
    local max_attempts=30
    local attempt=0

    while (( attempt < max_attempts )); do
        local status
        status=$(curl -s -o /dev/null -w "%{http_code}" "https://crates.io/api/v1/crates/${crate}/${version}")
        if [[ "$status" == "200" ]]; then
            log "  $crate $version is live on crates.io"
            return 0
        fi
        attempt=$((attempt + 1))
        echo "  Waiting for $crate $version to propagate... (attempt $attempt/$max_attempts)"
        sleep 10
    done

    fail "$crate $version did not appear on crates.io after $((max_attempts * 10))s"
}

publish_crate() {
    local crate="$1"
    log "Publishing $crate..."

    # Check if already published
    local status
    status=$(curl -s -o /dev/null -w "%{http_code}" "https://crates.io/api/v1/crates/${crate}/${VERSION}")
    if [[ "$status" == "200" ]]; then
        warn "$crate $VERSION already published, skipping."
        return 0
    fi

    run cargo publish -p "$crate" --allow-dirty
    if ! $DRY_RUN; then
        wait_for_crate "$crate" "$VERSION"
    fi
}

publish_level() {
    local level_name="$1"
    shift
    local crates=("$@")

    log "Step 2: Publishing $level_name to crates.io..."
    for crate in "${crates[@]}"; do
        publish_crate "$crate"
    done
    echo ""
}

if step_reached crates; then
    PUBLISH_DELAY=60  # seconds between levels to avoid crates.io rate limits

    publish_level "level 1 (core)"         "${CRATES_LEVEL_1[@]}"
    log "Waiting ${PUBLISH_DELAY}s for rate limit cooldown..."
    sleep "$PUBLISH_DELAY"
    publish_level "level 2 (markdown, docling)" "${CRATES_LEVEL_2[@]}"
    log "Waiting ${PUBLISH_DELAY}s for rate limit cooldown..."
    sleep "$PUBLISH_DELAY"
    publish_level "level 3 (format crates)" "${CRATES_LEVEL_3[@]}"
    log "Waiting ${PUBLISH_DELAY}s for rate limit cooldown..."
    sleep "$PUBLISH_DELAY"
    publish_level "level 4 (cli)"          "${CRATES_LEVEL_4[@]}"

    log "All Rust crates published."
    echo ""
fi

# ---------------------------------------------------------------------------
# Final verification
# ---------------------------------------------------------------------------

log "Final verification..."
echo ""

echo "Checking crates.io:"
ALL_CRATES=("${CRATES_LEVEL_1[@]}" "${CRATES_LEVEL_2[@]}" "${CRATES_LEVEL_3[@]}" "${CRATES_LEVEL_4[@]}")
for crate in "${ALL_CRATES[@]}"; do
    status=$(curl -s -o /dev/null -w "%{http_code}" "https://crates.io/api/v1/crates/${crate}/${VERSION}")
    if [[ "$status" == "200" ]]; then
        echo -e "  ${GREEN}ok${NC}  https://crates.io/crates/${crate}/${VERSION}"
    else
        echo -e "  ${RED}missing${NC}  https://crates.io/crates/${crate}/${VERSION}"
    fi
done

echo ""
log "crates.io release done."
log "PyPI and npm are published via CI: .github/workflows/release.yml"
log "  The tag push triggers the release workflow automatically."
