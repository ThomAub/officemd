# crates/tests spec contract

Tests here validate output stability across formats and bindings.

Contract:
- JSON snapshots are canonicalized (sorted keys, stable indentation).
- Markdown snapshots normalize newlines and trailing whitespace.
- Snapshot names map 1:1 to fixture and mode.
- Snapshot diffs require either a code fix or intentional refresh.

Checklist before merge:
- Run relevant snapshot tests and review diffs.
- Confirm fixture changes are intentional and documented.
- Ensure new tests fail before fix and pass after.
