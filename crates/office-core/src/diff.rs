//! The mandatory part-list-diff-on-save assertion.
//!
//! `office_oxide` has documented, open bugs where `save()` returns `Ok(())`
//! but silently drops whole zip parts (embedded OOXML packages, legacy VML
//! images, OLE objects — see the office-rs plan's ADR 0001). This module
//! converts that failure mode from silent to loud: snapshot every part path
//! in the document before an edit, and after `save()`, assert every part
//! that wasn't explicitly requested for removal is still present.
//!
//! Known gap, stated plainly rather than hidden: this only catches a part
//! *vanishing*. It does not catch content being silently truncated *inside*
//! a part that is still present under the same name (office_oxide issue
//! #230's failure shape). [`assert_no_suspicious_shrinkage`] is a cheap,
//! deliberately weaker secondary check for that gap: it does not prove
//! fidelity, it only catches gross truncation (a part that shrank to empty,
//! or valid XML that became non-well-formed).

use crate::error::{OfficeCoreError, Result};
use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::Read;
use std::path::Path;
use zip::ZipArchive;

/// A snapshot of one zip part, taken before a save, cheap enough to keep
/// one per part for every part in the document.
#[derive(Debug, Clone)]
pub struct PartSnapshot {
    pub size: u64,
    /// Only computed for parts whose name ends in `.xml` or `.rels` — true
    /// if the part's bytes parsed as well-formed XML at snapshot time.
    pub was_well_formed_xml: bool,
}

/// Every part in a document, keyed by its zip entry name (e.g.
/// `word/document.xml`, `word/embeddings/oleObject1.docx`).
pub type PartMap = HashMap<String, PartSnapshot>;

fn looks_like_xml(name: &str) -> bool {
    name.ends_with(".xml") || name.ends_with(".rels")
}

fn is_well_formed_xml(bytes: &[u8]) -> bool {
    quick_xml_check::Reader::from_bytes(bytes).check_well_formed()
}

/// Snapshot every part in the OOXML zip container at `path`.
///
/// This is deliberately a raw zip read with no office_oxide involvement —
/// it must not share office_oxide's own blind spots, since it exists to
/// catch office_oxide's bugs.
pub fn snapshot(path: &Path) -> Result<PartMap> {
    let file = File::open(path)?;
    let mut zip = ZipArchive::new(file)?;
    let mut parts = PartMap::new();
    for i in 0..zip.len() {
        let mut entry = zip.by_index(i)?;
        let name = entry.name().to_string();
        let size = entry.size();
        let was_well_formed_xml = if looks_like_xml(&name) {
            let mut buf = Vec::new();
            entry.read_to_end(&mut buf)?;
            is_well_formed_xml(&buf)
        } else {
            // Not XML — this check doesn't apply; treat as trivially true
            // so it never falsely fires for binary parts (images, OLE,
            // embedded packages).
            true
        };
        parts.insert(name, PartSnapshot { size, was_well_formed_xml });
    }
    Ok(parts)
}

/// The primary, hard assertion: every part present in `before` must still
/// be present in `after`, unless its name is in `explicitly_removed`.
///
/// Returns `Err(OfficeCoreError::PartLoss { .. })` naming every part that
/// vanished without being explicitly asked for.
pub fn assert_no_part_loss(
    before: &PartMap,
    after: &PartMap,
    explicitly_removed: &HashSet<String>,
) -> Result<()> {
    let missing: Vec<String> = before
        .keys()
        .filter(|name| !after.contains_key(*name) && !explicitly_removed.contains(*name))
        .cloned()
        .collect();
    if missing.is_empty() {
        Ok(())
    } else {
        Err(OfficeCoreError::PartLoss { missing_parts: missing })
    }
}

/// The secondary, weaker check for the stated gap: for every part present
/// in both snapshots, flag it if it shrank to zero bytes while it wasn't
/// zero before, or if it was well-formed XML before and isn't after.
///
/// This does not prove fidelity was preserved — a part can be rewritten
/// with real, different, valid content and this check says nothing about
/// that. It exists only to catch gross truncation matching office_oxide's
/// known #230 failure shape (84% content loss with no error).
pub fn assert_no_suspicious_shrinkage(before: &PartMap, after: &PartMap) -> Result<()> {
    let mut details = Vec::new();
    for (name, before_snap) in before {
        if let Some(after_snap) = after.get(name) {
            if before_snap.size > 0 && after_snap.size == 0 {
                details.push(format!(
                    "{name}: shrank to 0 bytes (was {} bytes)",
                    before_snap.size
                ));
            } else if before_snap.was_well_formed_xml && !after_snap.was_well_formed_xml {
                details.push(format!("{name}: was well-formed XML, no longer is"));
            }
        }
    }
    if details.is_empty() {
        Ok(())
    } else {
        Err(OfficeCoreError::SuspiciousPartChange { details })
    }
}

/// Convenience wrapper running both checks, in order, as every save path
/// in office-core must.
pub fn assert_save_preserved_fidelity(
    before: &PartMap,
    after: &PartMap,
    explicitly_removed: &HashSet<String>,
) -> Result<()> {
    assert_no_part_loss(before, after, explicitly_removed)?;
    assert_no_suspicious_shrinkage(before, after)?;
    Ok(())
}

/// Minimal, dependency-free well-formedness checker: a hand-rolled XML
/// tag-balance scanner. This deliberately avoids depending on a full XML
/// parser crate for office-core itself — it only needs to answer "is the
/// tag structure balanced and not truncated mid-tag," not parse semantics.
mod quick_xml_check {
    pub struct Reader<'a> {
        bytes: &'a [u8],
    }

    impl<'a> Reader<'a> {
        pub fn from_bytes(bytes: &'a [u8]) -> Self {
            Self { bytes }
        }

        /// Walks the byte stream tracking a stack of open tag names.
        /// Returns false on: an unclosed `<`, a mismatched close tag, or a
        /// non-empty stack at end-of-input. This is intentionally strict
        /// about well-formedness, not lenient about HTML-style tag soup —
        /// OOXML parts are always strict XML.
        pub fn check_well_formed(&self) -> bool {
            let s = match std::str::from_utf8(self.bytes) {
                Ok(s) => s,
                Err(_) => return false,
            };
            let mut stack: Vec<&str> = Vec::new();
            let mut chars = s.char_indices().peekable();
            while let Some((i, c)) = chars.next() {
                if c != '<' {
                    continue;
                }
                let rest = &s[i..];
                let close_idx = match rest.find('>') {
                    Some(idx) => idx,
                    None => return false, // unclosed tag
                };
                let tag_content = &rest[1..close_idx];
                // Skip the content we just scanned past.
                while let Some((j, _)) = chars.peek() {
                    if *j < i + close_idx + 1 {
                        chars.next();
                    } else {
                        break;
                    }
                }
                if tag_content.starts_with('?') || tag_content.starts_with('!') {
                    continue; // declaration / comment / doctype
                }
                if let Some(name) = tag_content.strip_prefix('/') {
                    let name = name.split_whitespace().next().unwrap_or("");
                    match stack.pop() {
                        Some(open) if open == name => {}
                        _ => return false, // mismatched or extra close tag
                    }
                } else if let Some(name) = tag_content.strip_suffix('/') {
                    let _ = name.split_whitespace().next().unwrap_or(""); // self-closing, no push
                } else {
                    let name = tag_content.split_whitespace().next().unwrap_or("");
                    if name.is_empty() {
                        return false;
                    }
                    stack.push(name);
                }
            }
            stack.is_empty()
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn snap(size: u64, well_formed: bool) -> PartSnapshot {
        PartSnapshot { size, was_well_formed_xml: well_formed }
    }

    #[test]
    fn no_loss_when_all_parts_survive() {
        let mut before = PartMap::new();
        before.insert("word/document.xml".into(), snap(100, true));
        let after = before.clone();
        assert!(assert_no_part_loss(&before, &after, &HashSet::new()).is_ok());
    }

    #[test]
    fn detects_silently_dropped_part() {
        let mut before = PartMap::new();
        before.insert("word/document.xml".into(), snap(100, true));
        before.insert("word/embeddings/oleObject1.docx".into(), snap(5000, true));
        let mut after = PartMap::new();
        after.insert("word/document.xml".into(), snap(100, true));
        let err = assert_no_part_loss(&before, &after, &HashSet::new()).unwrap_err();
        match err {
            OfficeCoreError::PartLoss { missing_parts } => {
                assert_eq!(missing_parts, vec!["word/embeddings/oleObject1.docx".to_string()]);
            }
            _ => panic!("expected PartLoss"),
        }
    }

    #[test]
    fn explicitly_removed_part_does_not_trigger() {
        let mut before = PartMap::new();
        before.insert("word/document.xml".into(), snap(100, true));
        before.insert("word/media/image1.png".into(), snap(2000, true));
        let mut after = PartMap::new();
        after.insert("word/document.xml".into(), snap(100, true));
        let mut removed = HashSet::new();
        removed.insert("word/media/image1.png".to_string());
        assert!(assert_no_part_loss(&before, &after, &removed).is_ok());
    }

    #[test]
    fn detects_shrink_to_zero() {
        let mut before = PartMap::new();
        before.insert("word/document.xml".into(), snap(5000, true));
        let mut after = PartMap::new();
        after.insert("word/document.xml".into(), snap(0, true));
        let err = assert_no_suspicious_shrinkage(&before, &after).unwrap_err();
        match err {
            OfficeCoreError::SuspiciousPartChange { details } => {
                assert_eq!(details.len(), 1);
                assert!(details[0].contains("word/document.xml"));
            }
            _ => panic!("expected SuspiciousPartChange"),
        }
    }

    #[test]
    fn detects_xml_becoming_malformed() {
        let mut before = PartMap::new();
        before.insert("word/document.xml".into(), snap(5000, true));
        let mut after = PartMap::new();
        after.insert("word/document.xml".into(), snap(1200, false));
        let err = assert_no_suspicious_shrinkage(&before, &after).unwrap_err();
        match err {
            OfficeCoreError::SuspiciousPartChange { details } => {
                assert!(details[0].contains("no longer is"));
            }
            _ => panic!("expected SuspiciousPartChange"),
        }
    }

    #[test]
    fn well_formed_checker_accepts_balanced_xml() {
        let xml = br#"<?xml version="1.0"?><root><a><b/></a><c>text</c></root>"#;
        assert!(is_well_formed_xml(xml));
    }

    #[test]
    fn well_formed_checker_rejects_unbalanced_xml() {
        let xml = br#"<?xml version="1.0"?><root><a><b></a></root>"#;
        assert!(!is_well_formed_xml(xml));
    }

    #[test]
    fn well_formed_checker_rejects_truncated_xml() {
        let xml = br#"<?xml version="1.0"?><root><a><b>trunc"#;
        assert!(!is_well_formed_xml(xml));
    }
}
