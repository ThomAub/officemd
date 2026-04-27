set shell := ["sh", "-cu"]

# List available commands
default:
    just --list

# Format Rust, Python, and JS sources
fmt:
    cargo fmt --all
    uv run --project crates/officemd_python ruff format crates/officemd_python examples/python
    cd crates/officemd_js && npm run format

# Check formatting for Rust, Python, and JS sources
fmt-check:
    cargo fmt --all --check
    uv run --project crates/officemd_python ruff format --check crates/officemd_python examples/python
    cd crates/officemd_js && npm run format:check

# Run linters for Rust, Python, and JS
lint:
    cargo clippy --workspace --exclude officemd_pdf --all-targets -- --force-warn clippy::pedantic -D warnings
    uv run --project crates/officemd_python ruff check crates/officemd_python examples/python
    uv run --project crates/officemd_python ty check crates/officemd_python/python examples/python
    cd crates/officemd_js && npm run lint

# Run the full local test suite
test: rust-test py-test js-test

# Run all Rust tests
rust-test:
    cargo test --workspace

# Run Rust check across the workspace
rust-check:
    cargo check --workspace

# Build/install the Python extension into the uv environment
py-develop:
    cd crates/officemd_python && uv run maturin develop --release

# Run Python tests
py-test:
    uv run --project crates/officemd_python pytest crates/officemd_python/tests crates/tests/snapshots -q

# Run Python type checks
py-typecheck:
    uv run --project crates/officemd_python ty check crates/officemd_python/python examples/python

# Install JS dependencies
js-install:
    cd crates/officemd_js && npm install

# Build JS native bindings
js-build:
    cd crates/officemd_js && npm run build

# Run JS tests
js-test:
    cd crates/officemd_js && node --test tests/cli-selection.test.mjs

# Regenerate Rust and Python markdown snapshots
snapshots-update:
    INSTA_UPDATE=always cargo test -p officemd_snapshot_tests
    cd crates/officemd_python && uv run maturin develop --release
    uv run --project crates/officemd_python pytest crates/tests/snapshots --force-regen -q

# Bump all package versions: just bump patch|minor|major
bump kind="patch":
    uv run bump.py {{ kind }}

# Preview a package version bump without writing changes
bump-dry-run kind="patch":
    uv run bump.py {{ kind }} --dry-run
