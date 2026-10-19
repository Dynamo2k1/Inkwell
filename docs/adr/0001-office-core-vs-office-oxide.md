# ADR 0001: office-core wraps office_oxide, with a mandatory part-list-diff assertion

## Status

Accepted. Implemented in Phase 0 (`office-core`, `office-docx`, `office-xlsx`, `office-pptx`,
`office-mcp`).

## Context

office-rs replaces `office-docs-mcp` (Python, python-docx/python-pptx), which covered 7 tools:
basic paragraph/table text for `.docx` and title+bullets slides for `.pptx`, no Excel at all. The
target is full read/write/edit coverage of `.docx`/`.xlsx`/`.pptx` with a mandatory fidelity
guarantee: never silently drop content on a round trip.

Building OOXML parsing/writing from scratch would duplicate real, substantial prior art.
`office_oxide` (crates.io, MIT/Apache-2.0, v0.1.15, github.com/yfedoseev/office_oxide) is an
actively maintained Rust library reading all six Office formats and writing/editing
`.docx`/`.xlsx`/`.pptx`, backed by a real fidelity-focused test corpus. But it has documented,
open, specific silent-content-loss bugs:

- Embedded OOXML packages in `.docx` silently dropped (issue #304).
- Legacy VML images silently dropped (#268).
- OLE objects in legacy `.doc` completely invisible (#284).
- A piece-table bug losing 84% of `.doc` content with no error (#230).
- `.doc`/`.xls`/`.ppt` metadata never read at all (#244).

Each of these is "save() returns `Ok(())`, but the content is gone" — not a crash, not a
reported error. That failure shape is exactly what office-rs exists to not repeat.

## Decision

`office-core` wraps `office_oxide` for the ordinary happy path (`office_core::opc::ReadHandle`
for reads, `office_core::opc::CreateSession`/`EditSession` for writes/edits), and enforces a
**mandatory part-list-diff assertion on every save**:

1. Snapshot every zip part path (and, for XML parts, byte size + well-formedness) before the
   operation.
2. Perform the office_oxide operation.
3. Snapshot the result the same way, via an independent raw-zip read with no office_oxide
   involvement (`office_core::diff::snapshot`) — it must not share office_oxide's own blind
   spots, since it exists to catch office_oxide's bugs.
4. Assert every part present before is still present after, unless explicitly marked removed
   (`office_core::diff::assert_no_part_loss`).
5. As a secondary, weaker check, assert no still-present part shrank to zero bytes or lost
   well-formed-XML status (`office_core::diff::assert_no_suspicious_shrinkage`).

This does not fix office_oxide's bugs. It converts every one of them from *silent* data loss
into a *loud, immediate* save-time error (`OfficeCoreError::PartLoss` /
`OfficeCoreError::SuspiciousPartChange`) — already a strict improvement over both office_oxide
alone and over python-docx/openpyxl today, neither of which has this discipline.

`office-docx`/`office-xlsx`/`office-pptx` depend only on `office-core`'s `opc` module, never on
`office_oxide` types directly in their own public signatures where avoidable, so the underlying
engine can be swapped later without reshaping their callers. (In practice `office_oxide`'s own
IR/format-specific types — `DocumentIR`, `CellValue`, `ImageFormat` — do appear in office-core's
and the format crates' public APIs, since re-wrapping every one of them would be a large surface
for no fidelity benefit; the boundary that matters is "never call `office_oxide::docx::write`
etc. directly outside `office-core`'s `opc` module," which holds.)

## Known, stated-plainly gap

The part-list diff only catches a part **vanishing**. If office_oxide silently truncates content
*inside* a part it still writes — e.g. dropping paragraphs from `document.xml` while still
emitting a valid, present `document.xml` (exactly issue #230's pattern: 84% content loss, no
error) — the part-list diff alone sees nothing wrong, because the part name is still there.

The secondary check (`assert_no_suspicious_shrinkage`) catches the grossest form of this: a part
shrinking to zero bytes, or XML that was well-formed becoming non-well-formed. It does **not**
prove fidelity was preserved — a part can be legitimately, correctly rewritten with real,
different content, and this check says nothing meaningful about that case. It exists only to
catch gross truncation matching the #230 failure shape, not to replace semantic diffing.

A real fix for the general case would require a semantic diff against the IR (counting
paragraphs/runs/table cells before and after an edit that should not have removed them) — not
built in Phase 0. If a future fidelity regression surfaces that this gap would have caught, that
is the trigger to build it, not a surprise.

## A second gap: office_oxide's edit API is narrower than its write API

While implementing `office-docx`/`office-pptx`, it became clear `office_oxide`'s *edit* types
(`EditableDocx`, `EditablePptx`) only support run-aware text find/replace on an already-existing
file — not structural insertion (new paragraphs, tables, images, slides). Structural writes to
an existing file are implemented here as **read-existing-IR, rebuild via the writer type
(`DocxWriter`/`PptxWriter`), append the new content, save to a (possibly new) path** — see
`rebuild_body_from_ir` in `office-docx` and `rebuild_slides_from_ir`/`apply_section_to_slide` in
`office-pptx`. This rebuild path only knows how to re-emit a subset of `Element` variants
(paragraphs, tables, images, lists for docx; title/paragraph-text/table for pptx) — any other
element kind present in the source document is preserved by the *read* side (so `docx_read`/
`pptx_read` never lose it) but **is dropped by a rebuild-based structural write**. This is
flagged inline in both crates' doc comments and in `office-mcp`'s tool descriptions, not hidden.

`office_oxide`'s XLSX edit type (`EditableXlsx`) is narrower still: it can only set individual
cell *values* on a sheet that already exists. It cannot add/rename/reorder sheets, set column
widths, or apply cell formatting to an existing file — those are only possible through
`XlsxWriter` at creation time. `office-xlsx::create` exposes that full creation-time surface;
`office-xlsx::write_cells` exposes only what `EditableXlsx` actually supports. **Freeze panes
have no write support anywhere in `office_oxide` v0.1.15, creation included** — there is no
`xlsx_freeze_panes` tool in `office-mcp` because there is nothing real for it to call.

## Consequences

- Every save path in office-rs is provably no-worse than silent — a part-list-diff failure is a
  hard error surfaced to the MCP caller, never swallowed.
- The MCP tool surface is honest about what it can and cannot do per format, rather than
  advertising operations that silently no-op (see `office-mcp/src/main.rs`'s module doc for the
  three tools from the original plan's table that do not exist, and why).
- Phase 1 (`office-legacy`, charts) and Phase 2 (live-automation tiers) build on `office-core`'s
  existing `opc`/`diff` modules without needing to change them — raw zip-entry passthrough
  (`office_core::opc::passthrough_raw_parts`) is already in place for Phase 1's known-buggy part
  classes (embedded packages, VML, OLE), it is just not yet wired into any crate's write path
  because Phase 0's fixtures didn't yet include a file exercising those bugs.
