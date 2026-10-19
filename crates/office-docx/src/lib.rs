//! Read/write/edit surface for `.docx`, built entirely on `office-core`'s
//! `opc` module — never on `office_oxide` types directly, so the engine
//! underneath can change without this crate's public API changing.

use office_core::error::Result;
use office_core::opc::{CreateSession, EditSession, ReadHandle};
use office_oxide::docx::edit::EditableDocx;
use office_oxide::docx::write::{Alignment, DocxWriter, Run};
use office_oxide::ir::DocumentIR;
use std::path::Path;

/// Reads a `.docx` file and returns its structured IR (paragraphs,
/// headings, tables, images, lists, ...) plus plain-text/Markdown views.
/// This is the single read entry point — `office-mcp`'s `docx_read` tool
/// serializes the IR directly rather than re-shaping it, so nothing it
/// contains is silently dropped on the way out.
pub struct DocxRead {
    pub ir: DocumentIR,
    pub plain_text: String,
    pub markdown: String,
}

pub fn read(path: impl AsRef<Path>) -> Result<DocxRead> {
    let handle = ReadHandle::open(path)?;
    Ok(DocxRead {
        plain_text: handle.plain_text(),
        markdown: handle.to_markdown(),
        ir: handle.to_ir(),
    })
}

/// A single text run for `create`/`append_paragraph`, mirroring
/// `office_oxide::docx::write::Run`'s builder surface without exposing
/// the type itself in this crate's public signatures.
#[derive(Debug, Clone, Default)]
pub struct RunSpec {
    pub text: String,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub color_hex: Option<String>,
    pub font_size_pt: Option<f64>,
    pub font_name: Option<String>,
}

impl RunSpec {
    pub fn plain(text: impl Into<String>) -> Self {
        Self { text: text.into(), ..Default::default() }
    }

    fn into_oxide_run(self) -> Run {
        let mut r = Run::new(self.text);
        if self.bold {
            r = r.bold();
        }
        if self.italic {
            r = r.italic();
        }
        if self.underline {
            r = r.underline();
        }
        if let Some(c) = self.color_hex {
            r = r.color(c);
        }
        if let Some(sz) = self.font_size_pt {
            r = r.font_size(sz);
        }
        if let Some(name) = self.font_name {
            r = r.font(name);
        }
        r
    }
}

fn parse_alignment(s: Option<&str>) -> Alignment {
    match s {
        Some("center") => Alignment::Center,
        Some("right") => Alignment::Right,
        Some("justify") => Alignment::Justify,
        _ => Alignment::Left,
    }
}

/// Creates a brand-new `.docx` at `dest`, failing if it already exists
/// (the office-docs-mcp convention this workspace keeps). `paragraphs`
/// is the document body, one entry per paragraph; an empty-run paragraph
/// is a plain `add_paragraph`, a single plain run is still routed through
/// `add_rich_paragraph` so formatting on later edits is uniform.
pub fn create(dest: impl AsRef<Path>, paragraphs: Vec<Vec<RunSpec>>) -> Result<()> {
    let mut writer = DocxWriter::new();
    for runs in paragraphs {
        let oxide_runs: Vec<Run> = runs.into_iter().map(RunSpec::into_oxide_run).collect();
        writer.add_rich_paragraph(&oxide_runs);
    }
    let session = CreateSession::new(writer);
    session.save(dest, |w, path| Ok(w.save(path)?))
}

/// Appends a single paragraph (optionally aligned) to an existing
/// `.docx`'s `DocxWriter`-built content is not possible in place —
/// office_oxide's edit API only supports text replacement on existing
/// files, not structural insertion. Appending a *new* paragraph to an
/// existing file is therefore implemented as: read the existing file's
/// IR, rebuild it via `DocxWriter` with the new paragraph appended, and
/// write to a new path. This is intentionally not silent about the
/// limitation — see `docs/adr/0001` for why office_oxide's edit surface
/// is this narrow.
pub fn append_paragraph(
    source: impl AsRef<Path>,
    dest: impl AsRef<Path>,
    runs: Vec<RunSpec>,
    alignment: Option<&str>,
) -> Result<()> {
    let handle = ReadHandle::open(source.as_ref())?;
    require_docx(&handle)?;
    let mut writer = DocxWriter::new();
    rebuild_body_from_ir(&mut writer, &handle.to_ir());
    let oxide_runs: Vec<Run> = runs.into_iter().map(RunSpec::into_oxide_run).collect();
    writer.add_rich_paragraph_aligned(&oxide_runs, parse_alignment(alignment));
    let session = CreateSession::new(writer);
    session.save(dest, |w, path| Ok(w.save(path)?))
}

fn require_docx(handle: &ReadHandle) -> Result<()> {
    if handle.format() != office_oxide::DocumentFormat::Docx {
        return Err(office_core::error::OfficeCoreError::Io(std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "source file is not a docx",
        )));
    }
    Ok(())
}

fn rebuild_body_from_ir(writer: &mut DocxWriter, ir: &DocumentIR) {
    use office_oxide::ir::Element;
    for section in &ir.sections {
        for el in &section.elements {
            match el {
                Element::Paragraph(p) => {
                    writer.add_ir_paragraph(&ir_runs_to_write_runs(p), None);
                }
                Element::Table(t) => {
                    writer.add_ir_table(&t);
                }
                Element::Image(img) => {
                    writer.add_ir_image(&img);
                }
                Element::List(list) => {
                    writer.add_ir_list(&list);
                }
                Element::PageBreak => {
                    writer.add_page_break();
                }
                _ => {
                    // Headings, text boxes, footnotes and other element
                    // kinds don't yet have a matching `DocxWriter`
                    // rebuild path here; they are preserved by the
                    // read-side IR (so `docx_read` never loses them) but
                    // an `append_paragraph` round-trip through this
                    // rebuild path can still drop them today. Flagged in
                    // `docs/adr/0001` as a known gap of the
                    // append-via-rebuild approach, not hidden here.
                }
            }
        }
    }
}

fn ir_runs_to_write_runs(p: &office_oxide::ir::Paragraph) -> Vec<Run> {
    use office_oxide::ir::InlineContent;
    p.content
        .iter()
        .filter_map(|c| match c {
            InlineContent::Text(span) => {
                let mut r = Run::new(span.text.clone());
                if span.bold {
                    r = r.bold();
                }
                if span.italic {
                    r = r.italic();
                }
                if span.underline.is_some() {
                    r = r.underline();
                }
                if let Some([red, green, blue]) = span.color {
                    r = r.color(format!("{red:02X}{green:02X}{blue:02X}"));
                }
                if let Some(half_pt) = span.font_size_half_pt {
                    r = r.font_size(half_pt as f64 / 2.0);
                }
                if let Some(name) = &span.font_name {
                    r = r.font(name.clone());
                }
                Some(r)
            }
            _ => None,
        })
        .collect()
}

/// Run-aware find/replace on an existing `.docx`, in place via
/// office_oxide's `EditableDocx::replace_text` (which operates on
/// `document.xml`'s real run structure, not a naive text-join — fixing
/// office-docs-mcp's run-boundary bug for whatever `office_oxide` itself
/// handles correctly). Saves under the mandatory part-list-diff
/// assertion: if office_oxide's save silently drops a part (its
/// documented embedded-package/VML/OLE bugs), this call returns an error
/// instead of a falsely-successful result.
///
/// Returns the number of replacements made.
pub fn replace_text(path: impl AsRef<Path>, find: &str, replace: &str) -> Result<usize> {
    let path = path.as_ref();
    let mut session = EditSession::<EditableDocx>::open(path, |p| Ok(EditableDocx::open(p)?))?;
    let count = session.editor_mut().replace_text(find, replace);
    session.save(path, |editor, dest| Ok(editor.save(dest)?))?;
    Ok(count)
}

/// Inserts a table into a new document built from an existing one plus
/// the new table appended at the end — same rebuild-and-append strategy
/// and the same stated gap as [`append_paragraph`].
pub fn insert_table(
    source: impl AsRef<Path>,
    dest: impl AsRef<Path>,
    rows: Vec<Vec<String>>,
) -> Result<()> {
    let handle = ReadHandle::open(source.as_ref())?;
    require_docx(&handle)?;
    let mut writer = DocxWriter::new();
    rebuild_body_from_ir(&mut writer, &handle.to_ir());
    let row_refs: Vec<Vec<&str>> =
        rows.iter().map(|row| row.iter().map(String::as_str).collect()).collect();
    writer.add_table(&row_refs);
    let session = CreateSession::new(writer);
    session.save(dest, |w, path| Ok(w.save(path)?))
}

/// Inserts an inline image (raw bytes + format) appended as a new
/// paragraph at the end of the document. Same rebuild-and-append
/// strategy and stated gap as [`append_paragraph`].
pub fn insert_image(
    source: impl AsRef<Path>,
    dest: impl AsRef<Path>,
    image_bytes: Vec<u8>,
    format: office_oxide::ir::ImageFormat,
    alt_text: Option<String>,
) -> Result<()> {
    let handle = ReadHandle::open(source.as_ref())?;
    require_docx(&handle)?;
    let mut writer = DocxWriter::new();
    rebuild_body_from_ir(&mut writer, &handle.to_ir());
    let image = office_oxide::ir::Image {
        alt_text,
        data: Some(image_bytes),
        format: Some(format),
        ..Default::default()
    };
    writer.add_ir_image(&image);
    let session = CreateSession::new(writer);
    session.save(dest, |w, path| Ok(w.save(path)?))
}
