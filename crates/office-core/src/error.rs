//! Shared error type for every format crate in this workspace.

use std::fmt;

/// Everything that can go wrong opening, editing, or saving an Office
/// document through this workspace — including `office_oxide`'s own
/// errors (wrapped, not hidden) and this workspace's own fidelity
/// guarantee tripping.
#[derive(Debug)]
pub enum OfficeCoreError {
    /// `office_oxide` itself returned an error (open/parse/save failure).
    OfficeOxide(office_oxide::OfficeError),
    /// Reading or writing the zip container for the part-list diff failed.
    Zip(String),
    /// The part-list diff assertion tripped: `save()` returned Ok, but one
    /// or more parts present before the edit are gone after it, and
    /// weren't explicitly requested for removal. This is the exact
    /// failure mode this workspace exists to catch loudly instead of
    /// letting it pass silently.
    PartLoss {
        /// Zip part paths (e.g. `word/embeddings/oleObject1.docx`) that
        /// were present before the save and are missing after.
        missing_parts: Vec<String>,
    },
    /// The cheaper secondary check: a part is still present by name, but
    /// its content looks grossly different (empty where it wasn't, or no
    /// longer well-formed XML where it was) — doesn't prove content is
    /// lost, but is worth surfacing rather than silently trusting.
    SuspiciousPartChange {
        /// The part path and a short description of what looked wrong.
        details: Vec<String>,
    },
    /// A plain I/O error unrelated to the zip/XML layers above.
    Io(std::io::Error),
}

impl fmt::Display for OfficeCoreError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::OfficeOxide(e) => write!(f, "office_oxide error: {e}"),
            Self::Zip(msg) => write!(f, "zip/part-list error: {msg}"),
            Self::PartLoss { missing_parts } => write!(
                f,
                "save() silently dropped {} part(s) that existed before and were not \
                 explicitly removed: {}",
                missing_parts.len(),
                missing_parts.join(", ")
            ),
            Self::SuspiciousPartChange { details } => write!(
                f,
                "save() left {} part(s) present but suspicious: {}",
                details.len(),
                details.join("; ")
            ),
            Self::Io(e) => write!(f, "I/O error: {e}"),
        }
    }
}

impl std::error::Error for OfficeCoreError {}

impl From<office_oxide::OfficeError> for OfficeCoreError {
    fn from(e: office_oxide::OfficeError) -> Self {
        Self::OfficeOxide(e)
    }
}

impl From<std::io::Error> for OfficeCoreError {
    fn from(e: std::io::Error) -> Self {
        Self::Io(e)
    }
}

impl From<zip::result::ZipError> for OfficeCoreError {
    fn from(e: zip::result::ZipError) -> Self {
        Self::Zip(e.to_string())
    }
}

pub type Result<T> = std::result::Result<T, OfficeCoreError>;
