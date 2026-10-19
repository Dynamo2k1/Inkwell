//! Read/create/edit surface for `.xlsx`, built entirely on `office-core`'s
//! `opc` module. See `docs/adr/0001` for the honest scope boundary this
//! crate keeps to: `office_oxide`'s XLSX *edit* API (`EditableXlsx`) only
//! supports setting individual cell values on an already-existing sheet —
//! it cannot add/rename/reorder sheets or set freeze panes on a file that
//! already exists. Those operations are only possible while *creating* a
//! new workbook with `XlsxWriter`. This is not an office-rs limitation
//! layered on top of a more capable engine — it is the real shape of
//! `office_oxide` v0.1.15's public API, and this crate does not pretend
//! otherwise by accepting parameters it can't honor.

use office_core::error::{OfficeCoreError, Result};
use office_core::opc::{CreateSession, EditSession, ReadHandle};
use office_oxide::ir::DocumentIR;
use office_oxide::xlsx::edit::{CellValue as EditCellValue, EditableXlsx};
use office_oxide::xlsx::write::{CellData, CellStyle, XlsxWriter};
use std::path::Path;

pub struct XlsxRead {
    pub ir: DocumentIR,
    pub plain_text: String,
    pub markdown: String,
}

pub fn read(path: impl AsRef<Path>) -> Result<XlsxRead> {
    let handle = ReadHandle::open(path)?;
    Ok(XlsxRead { plain_text: handle.plain_text(), markdown: handle.to_markdown(), ir: handle.to_ir() })
}

fn require_xlsx(handle: &ReadHandle) -> Result<()> {
    if handle.format() != office_oxide::DocumentFormat::Xlsx {
        return Err(OfficeCoreError::Io(std::io::Error::new(
            std::io::ErrorKind::InvalidInput,
            "source file is not an xlsx",
        )));
    }
    Ok(())
}

/// A single cell's value plus optional styling, for `create`'s sheet
/// data. Mirrors `office_oxide::xlsx::write::CellData`/`CellStyle`
/// without exposing those types in this crate's public signatures.
#[derive(Debug, Clone, Default)]
pub struct CellSpec {
    pub value: CellKind,
    pub bold: bool,
    pub italic: bool,
    pub font_color_hex: Option<String>,
    pub background_hex: Option<String>,
    pub number_format_code: Option<String>,
}

#[derive(Debug, Clone, Default)]
pub enum CellKind {
    #[default]
    Empty,
    Text(String),
    Number(f64),
    Boolean(bool),
}

impl CellSpec {
    pub fn text(s: impl Into<String>) -> Self {
        Self { value: CellKind::Text(s.into()), ..Default::default() }
    }

    pub fn number(n: f64) -> Self {
        Self { value: CellKind::Number(n), ..Default::default() }
    }

    fn has_style(&self) -> bool {
        self.bold
            || self.italic
            || self.font_color_hex.is_some()
            || self.background_hex.is_some()
            || self.number_format_code.is_some()
    }

    fn into_style(&self) -> CellStyle {
        let mut style = CellStyle::new();
        if self.bold {
            style = style.bold();
        }
        if self.italic {
            style = style.italic();
        }
        if let Some(c) = &self.font_color_hex {
            style = style.font_color(c.clone());
        }
        if let Some(c) = &self.background_hex {
            style = style.background(c.clone());
        }
        if let Some(code) = &self.number_format_code {
            style = style.number_format_code(code.clone());
        }
        style
    }

    fn into_cell_data(self) -> CellData {
        match self.value {
            CellKind::Empty => CellData::Empty,
            CellKind::Text(s) => CellData::String(s),
            CellKind::Number(n) => CellData::Number(n),
            CellKind::Boolean(b) => CellData::Boolean(b),
        }
    }
}

/// One sheet's worth of rows for `create`.
pub struct SheetSpec {
    pub name: String,
    pub rows: Vec<Vec<CellSpec>>,
}

/// Creates a brand-new `.xlsx` with one or more sheets, failing if `dest`
/// already exists. This is the only place multi-sheet authoring and
/// per-cell styling are available — see the module-level doc for why.
pub fn create(dest: impl AsRef<Path>, sheets: Vec<SheetSpec>) -> Result<()> {
    let mut writer = XlsxWriter::new();
    for sheet in sheets {
        let mut sheet_data = writer.add_sheet(&sheet.name);
        for (row_idx, row) in sheet.rows.into_iter().enumerate() {
            for (col_idx, cell) in row.into_iter().enumerate() {
                if cell.has_style() {
                    let style = cell.into_style();
                    let data = cell.into_cell_data();
                    sheet_data.set_cell_styled(row_idx, col_idx, data, style);
                } else {
                    let data = cell.into_cell_data();
                    sheet_data.set_cell(row_idx, col_idx, data);
                }
            }
        }
    }
    let session = CreateSession::new(writer);
    session.save(dest, |w, path| Ok(w.save(path)?))
}

/// Sets one or more individual cell values on an existing workbook's
/// already-existing sheets, in place, under the mandatory part-list-diff
/// assertion. `sheet_index` is 0-based per `office_oxide`'s own
/// convention. Returns the number of cells successfully set.
///
/// This cannot add a sheet, resize columns, or set freeze panes — see
/// the module-level doc. A cell write targeting a sheet index that
/// doesn't exist, or a chart/dialog sheet with no cells, fails loudly
/// (office_oxide's own documented behavior) rather than silently no-op.
pub fn write_cells(
    path: impl AsRef<Path>,
    writes: Vec<(usize, String, CellWriteValue)>,
) -> Result<usize> {
    let path = path.as_ref();
    let mut session = EditSession::<EditableXlsx>::open(path, |p| Ok(EditableXlsx::open(p)?))?;
    let mut count = 0;
    for (sheet_index, cell_ref, value) in writes {
        session
            .editor_mut()
            .set_cell(sheet_index, &cell_ref, value.into_edit_value())
            .map_err(office_oxide::OfficeError::from)?;
        count += 1;
    }
    session.save(path, |editor, dest| Ok(editor.save(dest)?))?;
    Ok(count)
}

#[derive(Debug, Clone)]
pub enum CellWriteValue {
    Empty,
    Text(String),
    Number(f64),
    Boolean(bool),
}

impl CellWriteValue {
    fn into_edit_value(self) -> EditCellValue {
        match self {
            Self::Empty => EditCellValue::Empty,
            Self::Text(s) => EditCellValue::String(s),
            Self::Number(n) => EditCellValue::Number(n),
            Self::Boolean(b) => EditCellValue::Boolean(b),
        }
    }
}

/// Validates `path` is an `.xlsx` before any read-side tool trusts it is
/// (format crates must call this explicitly — `office-core` stays
/// format-agnostic).
pub fn assert_is_xlsx(path: impl AsRef<Path>) -> Result<()> {
    let handle = ReadHandle::open(path)?;
    require_xlsx(&handle)
}
