# PDF visual text removal/replacement plan

Status: planning only. Do not implement in the OOXML formatting-preserving workstream.

## Goal

Provide a separate PDF-only feature for best-effort visual text removal/replacement.

This is **not** semantic PDF editing.
This is **not** guaranteed secure redaction.

## Non-goals

- Rewriting PDF text operators semantically in the general case
- Preserving editable text semantics
- Guaranteed exact font fidelity
- Security-grade redaction unless a dedicated redaction pass is designed and verified

## Proposed phases

### Phase 1: text-to-geometry mapping

Investigate current PDF extraction internals and confirm whether we can reliably obtain:
- page number
- text spans/words
- bounding boxes
- font size
- font name when available
- reading order

Questions to answer:
- how fragmented are words across spans/operators?
- can repeated words on a page be distinguished reliably?
- can we support rotated text?
- what happens on OCR'd vs scanned PDFs?

Deliverable:
- short feasibility report with examples and failure modes

### Phase 2: visual removal only

Implement best-effort text removal by:
1. locating matching text spans on a page
2. computing bounding boxes
3. painting opaque rectangles over those boxes

Suggested API shapes:

Rust:
```rust
PdfVisualReplace {
    from_text: "Confidential".into(),
    to_text: "".into(),
}
```

Python:
```python
remove_pdf_text(pdf_bytes, "Confidential")
```

Requirements:
- multiple matches per page
- multi-page support
- configurable fill color, default white
- clear failure/no-op behavior on scanned PDFs without reliable text geometry

### Phase 3: visual replacement overlay

After removal, optionally draw replacement text over the cleared region.

Requirements:
- reuse detected font size when possible
- best-effort font-family selection only
- graceful fallback font when original font cannot be reproduced
- preserve approximate placement, not exact kerning/ligatures

Suggested API shapes:

Rust:
```rust
PdfVisualReplace {
    from_text: "Old".into(),
    to_text: "New".into(),
}
```

Python:
```python
replace_pdf_text_visual(pdf_bytes, "Old", "New")
```

### Phase 4: redaction/security review

If product requirements later include secure redaction, this must be treated as a separate track.

Must investigate:
- whether original text remains in content streams
- incremental update behavior
- annotations and alternate representations
- object stream cleanup
- validation strategy for true redaction claims

## Test plan for PDF owner

Need fixtures for:
- simple selectable text PDFs
- repeated words on one page
- mixed fonts/sizes
- rotated text
- OCR'd PDFs
- scanned PDFs (expected no-op or explicit unsupported)
- multi-page documents

## User-facing caveats

The eventual docs should say:
- PDF replacement is visual/best effort
- OOXML formatting-preserving replacement is semantic and structure-aware
- PDF visual removal does not imply secure redaction
