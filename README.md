# Inkwell

**Read, create, and edit real Office files — docx, xlsx, pptx — without silently losing content.**

Inkwell is an MCP server (binary name `office-mcp`) that replaces `office-docs-mcp`. It wraps
[`office_oxide`](https://crates.io/crates/office_oxide) and adds one guarantee office_oxide alone
doesn't have: **every save is checked.** Before and after any edit, Inkwell snapshots the file's
real zip part list; if a save silently drops something — a known, documented failure mode in the
underlying library — that's a loud error, not a quiet corruption.

## What it actually does

Full read/create/edit support for `.docx`, `.xlsx`, `.pptx`. **Not** a universal Office file
handler — see "Known limitations" below before you rely on it for anything outside that.

| Tool | Does |
|---|---|
| `office_inspect` | Format, zip part list, core properties — the debug tool for a fidelity failure |
| `docx_read` / `docx_create` / `docx_append_paragraph` / `docx_replace_text` / `docx_insert_table` / `docx_insert_image` | Full docx surface |
| `xlsx_read` / `xlsx_create` / `xlsx_write_cells` | xlsx surface (see limitations — no sheet-add/formatting/freeze-panes on an existing file) |
| `pptx_read` / `pptx_create` / `pptx_add_slide` / `pptx_insert_image` / `pptx_insert_table` / `pptx_set_notes` | Full pptx surface |

All 16 tools are verified over real MCP stdio JSON-RPC (`initialize` → `tools/list` →
`tools/call`), not just at the Rust function level — including the three deliberate error paths
(`_create` refusing to overwrite, writing to a nonexistent sheet, setting notes on a nonexistent
slide).

## Known limitations (read before you trust this on an unfamiliar file)

- **Legacy binary formats (`.doc`/`.xls`/`.ppt`) are not supported.** The underlying library can
  technically read some of them, with documented bugs that silently lose up to 84% of a `.doc`
  file's content — not something to build on. No write/edit support exists for these at all.
  Fixing this (Phase 1) requires installing LibreOffice first, which hasn't been asked for or
  done yet.
- **xlsx**: `xlsx_write_cells` can only set a value in a cell that already exists on an already-
  existing sheet. It cannot add sheets, set formatting, or set freeze panes on an existing file —
  those are only possible via `xlsx_create` (a brand-new workbook). Freeze panes has **no** write
  support anywhere, creation included — the underlying library simply doesn't have it.
- **docx/pptx structural edits** (`docx_append_paragraph`, `docx_insert_table`,
  `docx_insert_image`, `pptx_add_slide`, `pptx_insert_table`, `pptx_insert_image`) work by
  rebuilding the whole document from its own structured representation and appending the new
  content — they do not patch the file in place. The rebuild only knows how to re-emit
  paragraphs/tables/images/lists (docx) or title/paragraph-text/tables (pptx). Reading a file
  (`docx_read`/`pptx_read`) always preserves everything; a *structural* edit can drop element
  kinds the rebuild path doesn't know about yet (footnotes, text boxes, complex layouts). Plain
  find/replace (`docx_replace_text`) edits in place and does not have this limitation.
  See `docs/adr/0001` for the full reasoning.
- Charts, macros/VBA, pivot table authoring, SmartArt, and animations are not supported for
  authoring in any mode.

## Architecture

```
crates/
  office-core/   office_oxide adapter + the mandatory part-list-diff fidelity assertion
  office-docx/   docx read/create/edit, built only on office-core
  office-xlsx/   xlsx read/create/edit, built only on office-core
  office-pptx/   pptx read/create/edit, built only on office-core
  office-mcp/    the rmcp stdio server (binary: office-mcp)
xtask/           cargo xtask fetch-fixtures / diff-report
tests/fixtures/  real, Office-authored fixtures (MIT-licensed), checksummed in MANIFEST.toml
docs/adr/        the two decisions worth reading before changing core behavior
```

See `docs/adr/0001-office-core-vs-office-oxide.md` for why this wraps office_oxide instead of
writing OOXML parsing from scratch, and the exact shape of the fidelity guarantee (and its
stated gap). See `docs/adr/0002-fixture-sourcing.md` for where test fixtures come from and why
none of the user's own real documents are in this repo.

## Build & test

```bash
cargo build --release -p office-mcp
cargo test --workspace
cargo run -p xtask -- fetch-fixtures      # verify fixture checksums
cargo run -p xtask -- diff-report <file>  # inspect a file's part list by hand
```

29 tests pass across the workspace, all against real fixtures (not synthetic bytes), including
explicit part-list-preservation checks on every edit path.

## Register with Claude Code

```bash
claude mcp add inkwell -s user -- /path/to/office-rs/target/release/office-mcp
```

## License

MIT.
