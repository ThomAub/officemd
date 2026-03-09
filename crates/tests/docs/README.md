# crates/tests docs scope

`crates/tests` contains shared regression coverage for crate outputs.

Current focus:
- snapshot tests for IR JSON and markdown generated from showcase fixtures
- cross-format guardrails to detect accidental output drift

This folder complements unit tests in individual crates by validating
workspace-level behavior against shared fixtures.
