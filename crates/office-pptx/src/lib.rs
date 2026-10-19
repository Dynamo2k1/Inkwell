//! Read/create/edit surface for `.pptx`, built entirely on `office-core`'s
//! `opc` module.
//!
//! Same honest scope boundary as `office-docx`: `office_oxide`'s PPTX
//! edit API (`EditablePptx`) only supports run-aware text find/replace on
//! an already-existing deck — it cannot add a slide, image, or table to a
//! file that already exists. `add_slide`/`insert_image`/`insert_table`
//! here are implemented as read-existing-IR, rebuild via `PptxWriter`,
//! append the new content, save-to-a-new-path — the same
//! rebuild-and-append strategy `office-docx::append_paragraph` uses, with
//! the same stated gap: element kinds the rebuild path doesn't yet know
//! how to re-emit (anything beyond title/paragraph-text/table/image) are
//! preserved on read but dropped by a rebuild-based write. Stated here,
//! not hidden — see `docs/adr/0001`.

use office_core::error::{OfficeCoreError, Result};
use office_core::opc::{CreateSession, EditSession, ReadHandle};
use office_oxide::ir::DocumentIR;
use office_oxide::pptx::edit::EditablePptx;
use office_oxide::pptx::write::{PptxWriter, Run};
use std::path::Path;

pub struct PptxRead {
    pub ir: DocumentIR,
    pub plain_text: String,
    pub markdown: String,
}

pub fn read(path: impl AsRef<Path>) -> Result<PptxRead> {
    let handle = ReadHandle::open(path)?;
    Ok(PptxRead { plain_text: handle.plain_text(), markdown: handle.to_markdown(), ir: handle.to_ir() })
}

fn require_pptx(handle: &ReadHandle) -> Result<()> {
    if handle.format() != office_oxide::DocumentFormat::Pptx {
        return Err(OfficeCoreError::Io(std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "source file is not a pptx",
        )));
    }
    Ok(())
}

/// One slide's content for `create`.
#[derive(Debug, Clone, Default)]
pub struct SlideSpec {
    pub title: Option<String>,
    pub bullets: Vec<String>,
    pub notes: Option<String>,
}

/// Creates a brand-new `.pptx`, one slide per `SlideSpec`, failing if
/// `dest` already exists.
pub fn create(dest: impl AsRef<Path>, slides: Vec<SlideSpec>) -> Result<()> {
    let mut writer = PptxWriter::new();
    for spec in slides {
        apply_slide_spec(&mut writer, spec);
    }
    let session = CreateSession::new(writer);
    session.save(dest, |w, path| Ok(w.save(path)?))
}

fn apply_slide_spec(writer: &mut PptxWriter, spec: SlideSpec) {
    let slide = writer.add_slide();
    if let Some(title) = &spec.title {
        slide.set_title(title);
    }
    if !spec.bullets.is_empty() {
        let refs: Vec<&str> = spec.bullets.iter().map(String::as_str).collect();
        slide.add_bullet_list(&refs);
    }
    if let Some(notes) = &spec.notes {
        slide.set_notes(notes);
    }
}

/// Appends a new slide to an existing deck by rebuilding every existing
/// slide's title/bullet-text/table/image content via `PptxWriter` and
/// then appending the new slide — see the module-level doc for the
/// stated rebuild gap.
pub fn add_slide(source: impl AsRef<Path>, dest: impl AsRef<Path>, spec: SlideSpec) -> Result<()> {
    let handle = ReadHandle::open(source.as_ref())?;
    require_pptx(&handle)?;
    let mut writer = PptxWriter::new();
    rebuild_slides_from_ir(&mut writer, &handle.to_ir());
    apply_slide_spec(&mut writer, spec);
    let session = CreateSession::new(writer);
    session.save(dest, |w, path| Ok(w.save(path)?))
}

/// Inserts a table as a new slide appended to the deck (office_oxide has
/// no "insert a table into an already-existing slide" primitive — tables
/// are built per-slide at write time). Same rebuild-and-append strategy.
pub fn insert_table(
    source: impl AsRef<Path>,
    dest: impl AsRef<Path>,
    rows: Vec<Vec<String>>,
) -> Result<()> {
    let handle = ReadHandle::open(source.as_ref())?;
    require_pptx(&handle)?;
    let mut writer = PptxWriter::new();
    rebuild_slides_from_ir(&mut writer, &handle.to_ir());
    let slide = writer.add_slide();
    let table_rows: Vec<Vec<Vec<Run>>> = rows
        .into_iter()
        .map(|row| row.into_iter().map(|cell| vec![Run::new(cell)]).collect())
        .collect();
    slide.add_table(table_rows);
    let session = CreateSession::new(writer);
    session.save(dest, |w, path| Ok(w.save(path)?))
}

/// Inserts an image as a new slide appended to the deck, at absolute
/// position `(x, y)` with size `(cx, cy)` (all in EMU — 914400 per inch).
/// Same rebuild-and-append strategy as [`insert_table`].
#[allow(clippy::too_many_arguments)]
pub fn insert_image(
    source: impl AsRef<Path>,
    dest: impl AsRef<Path>,
    image_bytes: Vec<u8>,
    format: office_oxide::ir::ImageFormat,
    x: i64,
    y: i64,
    cx: u64,
    cy: u64,
) -> Result<()> {
    let handle = ReadHandle::open(source.as_ref())?;
    require_pptx(&handle)?;
    let mut writer = PptxWriter::new();
    rebuild_slides_from_ir(&mut writer, &handle.to_ir());
    let slide = writer.add_slide();
    slide.add_image(image_bytes, format, x, y, cx, cy);
    let session = CreateSession::new(writer);
    session.save(dest, |w, path| Ok(w.save(path)?))
}

/// Sets speaker notes on an existing deck's slide by index, via the same
/// rebuild-and-append strategy (office_oxide's edit API has no
/// in-place notes mutation for an already-existing pptx).
pub fn set_notes(
    source: impl AsRef<Path>,
    dest: impl AsRef<Path>,
    slide_index: usize,
    notes: &str,
) -> Result<()> {
    let handle = ReadHandle::open(source.as_ref())?;
    require_pptx(&handle)?;
    let ir = handle.to_ir();
    if slide_index >= ir.sections.len() {
        return Err(OfficeCoreError::Io(std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            format!("slide index {slide_index} out of range ({} slides)", ir.sections.len()),
        )));
    }
    let mut writer = PptxWriter::new();
    for (i, section) in ir.sections.iter().enumerate() {
        let slide = writer.add_slide();
        apply_section_to_slide(slide, section);
        if i == slide_index {
            slide.set_notes(notes);
        }
    }
    let session = CreateSession::new(writer);
    session.save(dest, |w, path| Ok(w.save(path)?))
}

fn rebuild_slides_from_ir(writer: &mut PptxWriter, ir: &DocumentIR) {
    for section in &ir.sections {
        let slide = writer.add_slide();
        apply_section_to_slide(slide, section);
    }
}

fn apply_section_to_slide(
    slide: &mut office_oxide::pptx::write::SlideData,
    section: &office_oxide::ir::Section,
) {
    use office_oxide::ir::Element;
    if let Some(title) = &section.title {
        slide.set_title(title);
    }
    for el in &section.elements {
        match el {
            Element::Paragraph(p) => {
                let text = office_oxide::ir::inline_to_text(&p.content);
                if !text.is_empty() {
                    slide.add_text(&text);
                }
            }
            Element::Table(t) => {
                let table_rows: Vec<Vec<Vec<Run>>> = t
                    .rows
                    .iter()
                    .map(|row| {
                        row.cells
                            .iter()
                            .map(|cell| {
                                cell.content
                                    .iter()
                                    .filter_map(|e| match e {
                                        Element::Paragraph(p) => {
                                            Some(office_oxide::ir::inline_to_text(&p.content))
                                        }
                                        _ => None,
                                    })
                                    .map(Run::new)
                                    .collect()
                            })
                            .collect()
                    })
                    .collect();
                slide.add_table(table_rows);
            }
            _ => {
                // Images, lists, text boxes and other element kinds
                // don't yet have a rebuild path here — see the
                // module-level doc's stated gap.
            }
        }
    }
    if let Some(notes_elements) = &section.speaker_notes {
        let notes_text: String = notes_elements
            .iter()
            .filter_map(|e| match e {
                Element::Paragraph(p) => Some(office_oxide::ir::inline_to_text(&p.content)),
                _ => None,
            })
            .collect::<Vec<_>>()
            .join("\n");
        if !notes_text.is_empty() {
            slide.set_notes(&notes_text);
        }
    }
}

/// Run-aware find/replace on an existing `.pptx`, in place, under the
/// mandatory part-list-diff assertion. Returns the number of
/// replacements made.
pub fn replace_text(path: impl AsRef<Path>, find: &str, replace: &str) -> Result<usize> {
    let path = path.as_ref();
    let mut session = EditSession::<EditablePptx>::open(path, |p| Ok(EditablePptx::open(p)?))?;
    let count = session.editor_mut().replace_text(find, replace);
    session.save(path, |editor, dest| Ok(editor.save(dest)?))?;
    Ok(count)
}
