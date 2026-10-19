# ADR 0002: Fixture sourcing — real files only, consent-gated for the user's own documents

## Status

Accepted. Phase 0 fixtures vendored; `xtask fetch-fixtures` currently verifies local checksums
only (see "What's not built yet" below).

## Context

Round-trip fidelity testing against synthetic/hand-written OOXML would test office-rs against
office-rs's own assumptions about what real Office files look like — not against what Microsoft
Office, LibreOffice, or other real writers actually produce. The project's entire reason to exist
(catching office_oxide's silent-data-loss bugs) requires fixtures with the specific real-world
structure those bugs trigger on (embedded packages, VML, OLE, legacy piece tables) — synthetic
files built by us would not exercise them.

## Decision — sourcing priority

1. **Already on disk, zero cost**: `office-docs-mcp/.venv/.../docx/templates/default.docx` — a
   real, MIT-licensed python-docx template. Vendored as `tests/fixtures/docx_default_template.docx`.
2. **Other MIT-licensed open-source test corpora already present on this machine**: found via a
   filesystem search, `github.com/h2non/filetype`'s own test fixtures
   (`go/pkg/mod/github.com/h2non/filetype@v1.1.3/fixtures/`) include real, minimal, genuinely
   Office-authored `.xlsx` and `.pptx` files under MIT. Vendored as
   `tests/fixtures/xlsx_h2non_sample.xlsx` and `tests/fixtures/pptx_h2non_sample.pptx`. The same
   search also found real legacy `.doc`/`.xls`/`.ppt` fixtures in
   `github.com/gabriel-vasile/mimetype`'s test data (MIT) — not yet vendored, flagged in
   `tests/fixtures/MANIFEST.toml` as Phase 1 (`office-legacy`) candidates.
3. **Upstream sparse-clones** (python-docx/python-pptx/openpyxl's own `tests/` directories,
   Apache POI's `test-data` corpus) — described in the original plan, **not yet implemented**;
   `cargo xtask fetch-fixtures` today only verifies checksums of fixtures already vendored
   locally. Adding the actual network sparse-clone is future work, not required for Phase 0's
   own test suite to be meaningful (which it already is, against the three real fixtures above).
4. **The user's own real files** — a filesystem search during planning surfaced roughly 50 real
   Office files under `~/Downloads`, a mounted university drive, and various personal/business
   directories (`ZeroxInn/...`, `BugBounty/...`). These are personal, client, or business
   documents — several filenames on sight (a client vulnerability-findings report, a "sample
   report — redacted" document, a CV, device/network inventories, bug-bounty data) are reason
   enough not to decide unilaterally. **None have been vendored or read for fixture content.**
   Using any of them as a committed test fixture requires asking the user explicitly, file by
   file — this has not been done, because Phase 0 did not need to reach for personal files once
   categories 1–2 above provided real, safely-licensed material for docx/xlsx/pptx.

Fixtures are checksum-pinned in `tests/fixtures/MANIFEST.toml` (source, license, sha256) and
vendored directly (not fetched at test time) for Phase 0's three files, since they are tiny and
their provenance is fully documented. A future `fetch-fixtures` network implementation would
still record new fixtures the same way.

## What's not built yet

- The actual python-docx/python-pptx/openpyxl upstream sparse-clone and the Apache POI
  `test-data` corpus fetch — `xtask fetch-fixtures` currently only checksum-verifies what's
  already in `tests/fixtures/`.
- A "scrub" xtask subcommand (replace real text with placeholder text of equal length/
  formatting, keep structure) that would be the only path to committing a derivative of one of
  the user's own files — not built, because no consent conversation to use one has happened yet.

## Consequences

- Every Phase 0 test in `office-docx`/`office-xlsx`/`office-pptx` runs against at least one real,
  Office/library-authored file, not only against files office-rs itself created — `read_real_docx_template`
  and `read_real_pptx_fixture` specifically guard against office-rs only ever testing its own
  round-trip assumptions.
- No personal or client data has entered this repository's test fixtures at any point.
