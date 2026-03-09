# crates/tests agent guide

Scope: cross-crate regression tests and checked-in snapshots.

Rules:
- Favor deterministic assertions with canonicalized output.
- Keep tests independent of wall-clock time, locale, and network access.
- Reuse fixtures from `examples/data` unless a new shared fixture is required.
- Update snapshots only for intentional behavior changes.

When changing tests:
- Record why snapshot diffs are expected.
- Keep snapshot names stable and descriptive.
- Prefer focused tests over broad scripts.
