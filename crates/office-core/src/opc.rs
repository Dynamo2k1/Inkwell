//! Thin adapter over `office_oxide`, plus the save-time fidelity guarantee.
//!
//! Format crates (`office-docx`, `office-xlsx`, `office-pptx`) never call
//! `office_oxide` types directly — they go through [`EditSession`] and
//! [`ReadHandle`] here, so office_oxide can be swapped out later without
//! touching them, and so every save path gets the diff assertion for free
//! instead of each format crate having to remember to apply it.

use crate::diff::{self, PartMap};
use crate::error::Result;
use std::collections::HashSet;
use std::path::{Path, PathBuf};

/// A read-only handle on an existing document of any supported format.
/// Thin wrapper over `office_oxide::Document` — exists so format crates
/// import `office-core::opc::ReadHandle`, not `office_oxide::Document`.
pub struct ReadHandle {
    inner: office_oxide::Document,
}

impl ReadHandle {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Ok(Self { inner: office_oxide::Document::open(path)? })
    }

    pub fn plain_text(&self) -> String {
        self.inner.plain_text()
    }

    pub fn to_markdown(&self) -> String {
        self.inner.to_markdown()
    }

    pub fn to_ir(&self) -> office_oxide::ir::DocumentIR {
        self.inner.to_ir()
    }

    pub fn format(&self) -> office_oxide::DocumentFormat {
        self.inner.format()
    }

    pub fn as_docx(&self) -> Option<&office_oxide::docx::DocxDocument> {
        self.inner.as_docx()
    }

    pub fn as_xlsx(&self) -> Option<&office_oxide::xlsx::XlsxDocument> {
        self.inner.as_xlsx()
    }

    pub fn as_pptx(&self) -> Option<&office_oxide::pptx::PptxDocument> {
        self.inner.as_pptx()
    }
}

/// An in-progress edit of an existing OOXML file: snapshots the part list
/// at open time, hands the caller a mutation closure against office_oxide's
/// real editor type, and enforces [`diff::assert_save_preserved_fidelity`]
/// before reporting success — this is the single chokepoint every edit
/// path in this workspace goes through.
///
/// `E` is whichever office_oxide editor type applies (`EditableDocx`,
/// `EditableXlsx`, `EditablePptx`); `open_editor`/`do_save` are injected so
/// this struct stays format-agnostic.
pub struct EditSession<E> {
    source_path: PathBuf,
    before: PartMap,
    editor: E,
    explicitly_removed: HashSet<String>,
}

impl<E> EditSession<E> {
    /// Opens `path` for editing. `open_editor` must construct the
    /// office_oxide editor (e.g. `|p| EditableDocx::open(p)`) — injected
    /// rather than hardcoded so this one struct serves docx/xlsx/pptx.
    pub fn open(
        path: impl AsRef<Path>,
        open_editor: impl FnOnce(&Path) -> office_oxide::Result<E>,
    ) -> Result<Self> {
        let source_path = path.as_ref().to_path_buf();
        let before = diff::snapshot(&source_path)?;
        let editor = open_editor(&source_path)?;
        Ok(Self { source_path, before, editor, explicitly_removed: HashSet::new() })
    }

    /// Direct access to the office_oxide editor for format-specific
    /// mutation methods (`replace_text`, `set_cell`, ...).
    pub fn editor_mut(&mut self) -> &mut E {
        &mut self.editor
    }

    pub fn editor(&self) -> &E {
        &self.editor
    }

    /// Call before `save` when a mutation deliberately drops a part (e.g.
    /// deleting an image), so the diff assertion doesn't treat an
    /// intentional removal as a bug.
    pub fn mark_explicitly_removed(&mut self, part_name: impl Into<String>) {
        self.explicitly_removed.insert(part_name.into());
    }

    /// Saves to `dest` (which may equal the original path) via `do_save`,
    /// then runs the mandatory fidelity assertion against the freshly
    /// written file. On assertion failure, the write to `dest` has
    /// already happened — callers that need atomicity should save to a
    /// temp path and rename only after this returns `Ok`.
    pub fn save(
        &self,
        dest: impl AsRef<Path>,
        do_save: impl FnOnce(&E, &Path) -> office_oxide::Result<()>,
    ) -> Result<()> {
        let dest = dest.as_ref();
        do_save(&self.editor, dest)?;
        let after = diff::snapshot(dest)?;
        diff::assert_save_preserved_fidelity(&self.before, &after, &self.explicitly_removed)?;
        Ok(())
    }

    pub fn source_path(&self) -> &Path {
        &self.source_path
    }
}

/// A from-scratch document being built with one of office_oxide's writer
/// types (`DocxWriter`, `XlsxWriter`, pptx's slide-data API). There is no
/// "before" snapshot for creation — nothing existed to lose — so this is
/// a much thinner wrapper than [`EditSession`], mainly here so format
/// crates have one consistent `office-core::opc` entry point for both
/// create and edit instead of reaching into `office_oxide` directly for
/// one and `office-core` for the other.
pub struct CreateSession<W> {
    writer: W,
}

impl<W> CreateSession<W> {
    pub fn new(writer: W) -> Self {
        Self { writer }
    }

    pub fn writer_mut(&mut self) -> &mut W {
        &mut self.writer
    }

    pub fn writer(&self) -> &W {
        &self.writer
    }

    /// `_create` tools must fail if `dest` already exists — the
    /// office-docs-mcp convention this workspace keeps. Enforced here,
    /// once, rather than in every format crate.
    pub fn save(
        &self,
        dest: impl AsRef<Path>,
        do_save: impl FnOnce(&W, &Path) -> office_oxide::Result<()>,
    ) -> Result<()> {
        let dest = dest.as_ref();
        if dest.exists() {
            return Err(crate::error::OfficeCoreError::Io(std::io::Error::new(
                std::io::ErrorKind::AlreadyExists,
                format!("{} already exists; office-rs create tools never overwrite", dest.display()),
            )));
        }
        do_save(&self.writer, dest)?;
        Ok(())
    }
}

/// Raw zip-entry byte passthrough for office_oxide's known-buggy part
/// classes (embedded OOXML packages, legacy VML images, OLE objects) —
/// bypasses office_oxide's semantic model entirely for these parts rather
/// than trusting it to round-trip them. Copies the named entries from
/// `source` into `dest` byte-for-byte, after `dest` has already been
/// written by office_oxide's normal save path.
///
/// Returns the set of part names actually copied, suitable for passing as
/// context to callers deciding what to mark via
/// [`EditSession::mark_explicitly_removed`] (nothing here should be
/// removed — this function exists specifically to prevent removal).
pub fn passthrough_raw_parts(
    source: &Path,
    dest: &Path,
    part_names: &[&str],
) -> Result<HashSet<String>> {
    use std::io::{Read, Write};
    use zip::write::FileOptions;

    let src_file = std::fs::File::open(source)?;
    let mut src_zip = zip::ZipArchive::new(src_file)?;

    let mut raw_parts = HashSet::new();
    let mut buffers = Vec::new();
    for &name in part_names {
        if let Ok(mut entry) = src_zip.by_name(name) {
            let mut buf = Vec::new();
            entry.read_to_end(&mut buf)?;
            buffers.push((name.to_string(), buf));
            raw_parts.insert(name.to_string());
        }
    }
    drop(src_zip);

    if buffers.is_empty() {
        return Ok(raw_parts);
    }

    // Rewrite dest, replacing office_oxide's version of these parts (if
    // any) with the byte-exact original, by copying every other entry
    // through unchanged and appending the raw ones.
    let dest_bytes = std::fs::read(dest)?;
    let dest_reader = std::io::Cursor::new(dest_bytes);
    let mut dest_zip = zip::ZipArchive::new(dest_reader)?;

    let tmp_path = dest.with_extension("office-rs-tmp");
    let tmp_file = std::fs::File::create(&tmp_path)?;
    let mut writer = zip::ZipWriter::new(tmp_file);
    let replaced: HashSet<&str> = buffers.iter().map(|(n, _)| n.as_str()).collect();

    for i in 0..dest_zip.len() {
        let mut entry = dest_zip.by_index(i)?;
        let name = entry.name().to_string();
        if replaced.contains(name.as_str()) {
            continue; // will be written from `buffers` below instead
        }
        let options: FileOptions<()> = FileOptions::default();
        writer.start_file(&name, options)?;
        let mut buf = Vec::new();
        entry.read_to_end(&mut buf)?;
        writer.write_all(&buf)?;
    }
    for (name, buf) in &buffers {
        let options: FileOptions<()> = FileOptions::default();
        writer.start_file(name, options)?;
        writer.write_all(buf)?;
    }
    writer.finish()?;

    std::fs::rename(&tmp_path, dest)?;
    Ok(raw_parts)
}
