//! `office-mcp`: the MCP server for office-rs. Replaces `office-docs-mcp`.
//!
//! Tool surface and conventions follow the approved office-rs plan:
//! `_create` tools fail if the destination already exists (the
//! office-docs-mcp convention this workspace keeps); success responses
//! are a plain JSON object with a `path` field plus a summary; failures
//! surface office-core's real error text (including part-list-diff
//! fidelity failures) rather than a generic message.
//!
//! Three tools from the original plan's table are deliberately NOT
//! present: `xlsx_add_sheet`, `xlsx_set_formatting`, `xlsx_freeze_panes`
//! as *edit-an-existing-file* operations. Verified against
//! `office_oxide` v0.1.15's real public API while implementing
//! `office-xlsx`: its `EditableXlsx` type only supports setting
//! individual cell values on a sheet that already exists — adding
//! sheets, setting column/cell formatting, and freeze panes are only
//! possible while *creating* a new workbook via `XlsxWriter`, which
//! `xlsx_create` already exposes in full. Freeze panes specifically has
//! no write support anywhere in `office_oxide` at all, creation
//! included. Advertising tools that silently no-op or error on every
//! call would be the exact kind of dishonesty this project exists to
//! avoid, so they are omitted rather than stubbed.

use anyhow::anyhow;
use base64::Engine;
use base64::engine::general_purpose::STANDARD as BASE64;
use rmcp::handler::server::router::tool::ToolRouter;
use rmcp::handler::server::wrapper::Parameters;
use rmcp::model::{CallToolResult, ContentBlock, Implementation, ServerCapabilities, ServerConfig};
use rmcp::transport::stdio;
use rmcp::{ServerHandler, ServiceExt, tool, tool_handler, tool_router};
use schemars::JsonSchema;
use serde::Deserialize;
use serde_json::json;

fn ok_json(value: serde_json::Value) -> CallToolResult {
    CallToolResult::success(vec![ContentBlock::text(value.to_string())])
}

fn error_result(message: impl std::fmt::Display) -> CallToolResult {
    CallToolResult::error(vec![ContentBlock::text(json!({ "error": message.to_string() }).to_string())])
}

fn parse_image_format(s: &str) -> Result<office_oxide::ir::ImageFormat, String> {
    use office_oxide::ir::ImageFormat::*;
    match s.to_ascii_lowercase().as_str() {
        "png" => Ok(Png),
        "jpeg" | "jpg" => Ok(Jpeg),
        "gif" => Ok(Gif),
        "tiff" => Ok(Tiff),
        "bmp" => Ok(Bmp),
        "emf" => Ok(Emf),
        "wmf" => Ok(Wmf),
        other => Err(format!("unsupported image format: {other} (expected png/jpeg/gif/tiff/bmp/emf/wmf)")),
    }
}

// ---------------------------------------------------------------------
// Shared / inspect
// ---------------------------------------------------------------------

#[derive(Debug, JsonSchema, Deserialize)]
struct PathArgs {
    /// Absolute or relative path to the Office file.
    path: String,
}

#[derive(Debug, JsonSchema, Deserialize)]
struct RunArg {
    text: String,
    #[serde(default)]
    bold: bool,
    #[serde(default)]
    italic: bool,
    #[serde(default)]
    underline: bool,
    /// Hex RGB, e.g. "FF0000".
    color_hex: Option<String>,
    font_size_pt: Option<f64>,
    font_name: Option<String>,
}

#[derive(Debug, JsonSchema, Deserialize)]
struct DocxCreateArgs {
    path: String,
    /// One entry per paragraph; each paragraph is a list of runs.
    paragraphs: Vec<Vec<RunArg>>,
}

#[derive(Debug, JsonSchema, Deserialize)]
struct DocxAppendParagraphArgs {
    source: String,
    dest: String,
    runs: Vec<RunArg>,
    /// left, center, right, or justify (default left).
    alignment: Option<String>,
}

#[derive(Debug, JsonSchema, Deserialize)]
struct ReplaceTextArgs {
    path: String,
    find: String,
    replace: String,
}

#[derive(Debug, JsonSchema, Deserialize)]
struct InsertTableArgs {
    source: String,
    dest: String,
    /// Row-major table cells (plain text).
    rows: Vec<Vec<String>>,
}

#[derive(Debug, JsonSchema, Deserialize)]
struct DocxInsertImageArgs {
    source: String,
    dest: String,
    /// Base64-encoded raw image bytes.
    image_base64: String,
    /// png/jpeg/gif/tiff/bmp/emf/wmf.
    format: String,
    alt_text: Option<String>,
}

#[derive(Debug, JsonSchema, Deserialize)]
struct XlsxCellArg {
    #[serde(default)]
    text: Option<String>,
    #[serde(default)]
    number: Option<f64>,
    #[serde(default)]
    boolean: Option<bool>,
    #[serde(default)]
    bold: bool,
    #[serde(default)]
    italic: bool,
    font_color_hex: Option<String>,
    background_hex: Option<String>,
    number_format_code: Option<String>,
}

impl XlsxCellArg {
    fn into_spec(self) -> office_xlsx::CellSpec {
        let value = if let Some(s) = self.text {
            office_xlsx::CellKind::Text(s)
        } else if let Some(n) = self.number {
            office_xlsx::CellKind::Number(n)
        } else if let Some(b) = self.boolean {
            office_xlsx::CellKind::Boolean(b)
        } else {
            office_xlsx::CellKind::Empty
        };
        office_xlsx::CellSpec {
            value,
            bold: self.bold,
            italic: self.italic,
            font_color_hex: self.font_color_hex,
            background_hex: self.background_hex,
            number_format_code: self.number_format_code,
        }
    }
}

#[derive(Debug, JsonSchema, Deserialize)]
struct XlsxSheetArg {
    name: String,
    rows: Vec<Vec<XlsxCellArg>>,
}

#[derive(Debug, JsonSchema, Deserialize)]
struct XlsxCreateArgs {
    path: String,
    sheets: Vec<XlsxSheetArg>,
}

#[derive(Debug, JsonSchema, Deserialize)]
struct XlsxCellWrite {
    sheet_index: usize,
    /// e.g. "A1".
    cell_ref: String,
    #[serde(default)]
    text: Option<String>,
    #[serde(default)]
    number: Option<f64>,
    #[serde(default)]
    boolean: Option<bool>,
}

impl XlsxCellWrite {
    fn into_write(self) -> (usize, String, office_xlsx::CellWriteValue) {
        let value = if let Some(s) = self.text {
            office_xlsx::CellWriteValue::Text(s)
        } else if let Some(n) = self.number {
            office_xlsx::CellWriteValue::Number(n)
        } else if let Some(b) = self.boolean {
            office_xlsx::CellWriteValue::Boolean(b)
        } else {
            office_xlsx::CellWriteValue::Empty
        };
        (self.sheet_index, self.cell_ref, value)
    }
}

#[derive(Debug, JsonSchema, Deserialize)]
struct XlsxWriteCellsArgs {
    path: String,
    writes: Vec<XlsxCellWrite>,
}

#[derive(Debug, JsonSchema, Deserialize)]
struct PptxSlideArg {
    title: Option<String>,
    #[serde(default)]
    bullets: Vec<String>,
    notes: Option<String>,
}

impl PptxSlideArg {
    fn into_spec(self) -> office_pptx::SlideSpec {
        office_pptx::SlideSpec { title: self.title, bullets: self.bullets, notes: self.notes }
    }
}

#[derive(Debug, JsonSchema, Deserialize)]
struct PptxCreateArgs {
    path: String,
    slides: Vec<PptxSlideArg>,
}

#[derive(Debug, JsonSchema, Deserialize)]
struct PptxAddSlideArgs {
    source: String,
    dest: String,
    slide: PptxSlideArg,
}

#[derive(Debug, JsonSchema, Deserialize)]
struct PptxInsertImageArgs {
    source: String,
    dest: String,
    image_base64: String,
    format: String,
    /// EMU (914400 per inch). Defaults position the image at the
    /// top-left in a 1-inch box if omitted.
    x: Option<i64>,
    y: Option<i64>,
    cx: Option<u64>,
    cy: Option<u64>,
}

#[derive(Debug, JsonSchema, Deserialize)]
struct PptxSetNotesArgs {
    source: String,
    dest: String,
    slide_index: usize,
    notes: String,
}

#[derive(Clone)]
struct OfficeMcp {
    tool_router: ToolRouter<OfficeMcp>,
}

#[tool_router]
impl OfficeMcp {
    fn new() -> Self {
        Self { tool_router: Self::tool_router() }
    }

    #[tool(description = "Inspect any supported Office file: format, zip part list, core properties. Also the debug tool for understanding a part-list-diff fidelity failure.")]
    async fn office_inspect(&self, Parameters(args): Parameters<PathArgs>) -> CallToolResult {
        match office_core::opc::ReadHandle::open(&args.path) {
            Ok(handle) => {
                let ir = handle.to_ir();
                let parts = match office_core::diff::snapshot(std::path::Path::new(&args.path)) {
                    Ok(p) => {
                        let mut names: Vec<String> = p.keys().cloned().collect();
                        names.sort();
                        names
                    }
                    Err(_) => Vec::new(),
                };
                ok_json(json!({
                    "path": args.path,
                    "format": format!("{:?}", handle.format()),
                    "part_count": parts.len(),
                    "parts": parts,
                    "metadata": ir.metadata,
                    "section_count": ir.sections.len(),
                }))
            }
            Err(e) => error_result(e),
        }
    }

    // --- docx ---

    #[tool(description = "Read a .docx file: returns its structured IR (paragraphs, headings, tables, images, lists, ...), plain text, and markdown.")]
    async fn docx_read(&self, Parameters(args): Parameters<PathArgs>) -> CallToolResult {
        match office_docx::read(&args.path) {
            Ok(r) => ok_json(json!({ "ir": r.ir, "plain_text": r.plain_text, "markdown": r.markdown })),
            Err(e) => error_result(e),
        }
    }

    #[tool(description = "Create a new .docx. Fails if the path already exists.")]
    async fn docx_create(&self, Parameters(args): Parameters<DocxCreateArgs>) -> CallToolResult {
        let paragraphs: Vec<Vec<office_docx::RunSpec>> = args
            .paragraphs
            .into_iter()
            .map(|p| p.into_iter().map(run_arg_to_spec).collect())
            .collect();
        match office_docx::create(&args.path, paragraphs) {
            Ok(()) => ok_json(json!({ "path": args.path, "created": true })),
            Err(e) => error_result(e),
        }
    }

    #[tool(description = "Append a new paragraph to an existing .docx, writing the result to `dest` (may differ from `source`). Rebuilds the document from its own IR, so element kinds without a rebuild path (beyond paragraphs/tables/images/lists) may be dropped — see office-rs docs/adr/0001.")]
    async fn docx_append_paragraph(&self, Parameters(args): Parameters<DocxAppendParagraphArgs>) -> CallToolResult {
        let runs: Vec<office_docx::RunSpec> = args.runs.into_iter().map(run_arg_to_spec).collect();
        match office_docx::append_paragraph(&args.source, &args.dest, runs, args.alignment.as_deref()) {
            Ok(()) => ok_json(json!({ "path": args.dest, "appended": true })),
            Err(e) => error_result(e),
        }
    }

    #[tool(description = "Run-aware find/replace on an existing .docx, in place, under office-rs's mandatory part-list-diff fidelity check.")]
    async fn docx_replace_text(&self, Parameters(args): Parameters<ReplaceTextArgs>) -> CallToolResult {
        match office_docx::replace_text(&args.path, &args.find, &args.replace) {
            Ok(count) => ok_json(json!({ "path": args.path, "replacements": count })),
            Err(e) => error_result(e),
        }
    }

    #[tool(description = "Append a table to an existing .docx, writing the result to `dest`.")]
    async fn docx_insert_table(&self, Parameters(args): Parameters<InsertTableArgs>) -> CallToolResult {
        match office_docx::insert_table(&args.source, &args.dest, args.rows) {
            Ok(()) => ok_json(json!({ "path": args.dest, "inserted": true })),
            Err(e) => error_result(e),
        }
    }

    #[tool(description = "Append an inline image (base64-encoded bytes) to an existing .docx, writing the result to `dest`.")]
    async fn docx_insert_image(&self, Parameters(args): Parameters<DocxInsertImageArgs>) -> CallToolResult {
        let format = match parse_image_format(&args.format) {
            Ok(f) => f,
            Err(e) => return error_result(e),
        };
        let bytes = match BASE64.decode(&args.image_base64) {
            Ok(b) => b,
            Err(e) => return error_result(format!("invalid base64: {e}")),
        };
        match office_docx::insert_image(&args.source, &args.dest, bytes, format, args.alt_text) {
            Ok(()) => ok_json(json!({ "path": args.dest, "inserted": true })),
            Err(e) => error_result(e),
        }
    }

    // --- xlsx ---

    #[tool(description = "Read an .xlsx file: returns its structured IR (one section per sheet), plain text, and markdown.")]
    async fn xlsx_read(&self, Parameters(args): Parameters<PathArgs>) -> CallToolResult {
        match office_xlsx::read(&args.path) {
            Ok(r) => ok_json(json!({ "ir": r.ir, "plain_text": r.plain_text, "markdown": r.markdown })),
            Err(e) => error_result(e),
        }
    }

    #[tool(description = "Create a new .xlsx with one or more sheets, each with cell values and optional per-cell styling. Fails if the path already exists. This is the only office-rs tool that can add multiple sheets or set per-cell formatting — office_oxide's edit API for an already-existing workbook only supports setting individual cell values (see xlsx_write_cells); it has no sheet-add, formatting, or freeze-pane support at all for existing files, and no freeze-pane support in any mode.")]
    async fn xlsx_create(&self, Parameters(args): Parameters<XlsxCreateArgs>) -> CallToolResult {
        let sheets: Vec<office_xlsx::SheetSpec> = args
            .sheets
            .into_iter()
            .map(|s| office_xlsx::SheetSpec {
                name: s.name,
                rows: s.rows.into_iter().map(|row| row.into_iter().map(XlsxCellArg::into_spec).collect()).collect(),
            })
            .collect();
        match office_xlsx::create(&args.path, sheets) {
            Ok(()) => ok_json(json!({ "path": args.path, "created": true })),
            Err(e) => error_result(e),
        }
    }

    #[tool(description = "Set one or more individual cell values on an existing .xlsx's already-existing sheets, in place, under the mandatory part-list-diff fidelity check. Cannot add sheets or set formatting — use xlsx_create for that.")]
    async fn xlsx_write_cells(&self, Parameters(args): Parameters<XlsxWriteCellsArgs>) -> CallToolResult {
        let writes = args.writes.into_iter().map(XlsxCellWrite::into_write).collect();
        match office_xlsx::write_cells(&args.path, writes) {
            Ok(count) => ok_json(json!({ "path": args.path, "cells_written": count })),
            Err(e) => error_result(e),
        }
    }

    // --- pptx ---

    #[tool(description = "Read a .pptx file: returns its structured IR (one section per slide), plain text, and markdown.")]
    async fn pptx_read(&self, Parameters(args): Parameters<PathArgs>) -> CallToolResult {
        match office_pptx::read(&args.path) {
            Ok(r) => ok_json(json!({ "ir": r.ir, "plain_text": r.plain_text, "markdown": r.markdown })),
            Err(e) => error_result(e),
        }
    }

    #[tool(description = "Create a new .pptx, one slide per entry (title, bullet list, speaker notes). Fails if the path already exists.")]
    async fn pptx_create(&self, Parameters(args): Parameters<PptxCreateArgs>) -> CallToolResult {
        let slides: Vec<office_pptx::SlideSpec> = args.slides.into_iter().map(PptxSlideArg::into_spec).collect();
        match office_pptx::create(&args.path, slides) {
            Ok(()) => ok_json(json!({ "path": args.path, "created": true })),
            Err(e) => error_result(e),
        }
    }

    #[tool(description = "Append a new slide to an existing .pptx, writing the result to `dest`. Rebuilds the deck from its own IR; element kinds beyond title/paragraph-text/table/image may be dropped on rebuild — see office-rs docs/adr/0001.")]
    async fn pptx_add_slide(&self, Parameters(args): Parameters<PptxAddSlideArgs>) -> CallToolResult {
        match office_pptx::add_slide(&args.source, &args.dest, args.slide.into_spec()) {
            Ok(()) => ok_json(json!({ "path": args.dest, "appended": true })),
            Err(e) => error_result(e),
        }
    }

    #[tool(description = "Append an image (base64-encoded bytes) as a new slide to an existing .pptx, writing the result to `dest`.")]
    async fn pptx_insert_image(&self, Parameters(args): Parameters<PptxInsertImageArgs>) -> CallToolResult {
        let format = match parse_image_format(&args.format) {
            Ok(f) => f,
            Err(e) => return error_result(e),
        };
        let bytes = match BASE64.decode(&args.image_base64) {
            Ok(b) => b,
            Err(e) => return error_result(format!("invalid base64: {e}")),
        };
        let (x, y, cx, cy) = (args.x.unwrap_or(0), args.y.unwrap_or(0), args.cx.unwrap_or(914_400), args.cy.unwrap_or(914_400));
        match office_pptx::insert_image(&args.source, &args.dest, bytes, format, x, y, cx, cy) {
            Ok(()) => ok_json(json!({ "path": args.dest, "inserted": true })),
            Err(e) => error_result(e),
        }
    }

    #[tool(description = "Append a table as a new slide to an existing .pptx, writing the result to `dest`.")]
    async fn pptx_insert_table(&self, Parameters(args): Parameters<InsertTableArgs>) -> CallToolResult {
        match office_pptx::insert_table(&args.source, &args.dest, args.rows) {
            Ok(()) => ok_json(json!({ "path": args.dest, "inserted": true })),
            Err(e) => error_result(e),
        }
    }

    #[tool(description = "Set speaker notes on an existing .pptx slide by 0-based index, writing the result to `dest`.")]
    async fn pptx_set_notes(&self, Parameters(args): Parameters<PptxSetNotesArgs>) -> CallToolResult {
        match office_pptx::set_notes(&args.source, &args.dest, args.slide_index, &args.notes) {
            Ok(()) => ok_json(json!({ "path": args.dest, "notes_set": true })),
            Err(e) => error_result(e),
        }
    }
}

fn run_arg_to_spec(r: RunArg) -> office_docx::RunSpec {
    office_docx::RunSpec {
        text: r.text,
        bold: r.bold,
        italic: r.italic,
        underline: r.underline,
        color_hex: r.color_hex,
        font_size_pt: r.font_size_pt,
        font_name: r.font_name,
    }
}

#[tool_handler]
impl ServerHandler for OfficeMcp {
    fn get_info(&self) -> ServerConfig {
        ServerConfig::new(ServerCapabilities::builder().enable_tools().build())
            .with_server_info(Implementation::new("office-mcp", env!("CARGO_PKG_VERSION")))
    }
}

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    let service = OfficeMcp::new()
        .serve(stdio())
        .await
        .map_err(|e| anyhow!("failed to start MCP server: {e}"))?;
    service.waiting().await?;
    Ok(())
}
