//! Workspace maintenance tasks for office-rs.
//!
//! `cargo xtask fetch-fixtures` — verifies every fixture listed in
//! `tests/fixtures/MANIFEST.toml` is present and matches its recorded
//! sha256, catching silent drift between the manifest and what's on
//! disk. It deliberately does NOT download anything from the network
//! yet (the python-docx/python-pptx/openpyxl upstream sparse-clone and
//! the Apache POI test-data corpus fetch described in the office-rs plan
//! are future work, not yet implemented) — it only verifies the fixtures
//! already vendored locally, so running it is still meaningful today.
//!
//! `cargo xtask diff-report <file>` — runs the same part-list snapshot
//! office-core's save-time assertion uses, and prints it, so a fidelity
//! failure can be inspected by hand without writing a one-off script.

use anyhow::{Context, Result, bail};
use sha2::{Digest, Sha256};
use std::path::{Path, PathBuf};

fn workspace_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).parent().unwrap().to_path_buf()
}

fn main() -> Result<()> {
    let args: Vec<String> = std::env::args().skip(1).collect();
    match args.first().map(String::as_str) {
        Some("fetch-fixtures") => fetch_fixtures(),
        Some("diff-report") => {
            let path = args.get(1).context("usage: cargo xtask diff-report <file>")?;
            diff_report(Path::new(path))
        }
        _ => {
            eprintln!("usage: cargo xtask <fetch-fixtures|diff-report <file>>");
            std::process::exit(1);
        }
    }
}

#[derive(serde::Deserialize)]
struct Manifest {
    fixture: Vec<FixtureEntry>,
}

#[derive(serde::Deserialize)]
struct FixtureEntry {
    file: String,
    source: String,
    sha256: String,
}

fn fetch_fixtures() -> Result<()> {
    let manifest_path = workspace_root().join("tests/fixtures/MANIFEST.toml");
    let manifest_text = std::fs::read_to_string(&manifest_path)
        .with_context(|| format!("reading {}", manifest_path.display()))?;
    let manifest: Manifest = parse_manifest(&manifest_text)?;

    let fixtures_dir = workspace_root().join("tests/fixtures");
    let mut ok = 0;
    let mut failed = Vec::new();
    for entry in &manifest.fixture {
        let path = fixtures_dir.join(&entry.file);
        match verify_checksum(&path, &entry.sha256) {
            Ok(()) => {
                ok += 1;
                println!("OK   {} ({})", entry.file, entry.source);
            }
            Err(e) => {
                failed.push(entry.file.clone());
                eprintln!("FAIL {} — {e}", entry.file);
            }
        }
    }
    println!("{ok}/{} fixtures verified", manifest.fixture.len());
    if !failed.is_empty() {
        bail!("fixture checksum mismatch or missing file: {}", failed.join(", "));
    }
    Ok(())
}

fn verify_checksum(path: &Path, expected_sha256: &str) -> Result<()> {
    let data = std::fs::read(path).with_context(|| format!("reading {}", path.display()))?;
    let mut hasher = Sha256::new();
    hasher.update(&data);
    let actual = hex_encode(&hasher.finalize());
    if actual != expected_sha256.to_ascii_lowercase() {
        bail!("checksum mismatch: expected {expected_sha256}, got {actual}");
    }
    Ok(())
}

fn hex_encode(bytes: &[u8]) -> String {
    bytes.iter().map(|b| format!("{b:02x}")).collect()
}

/// Minimal hand-rolled TOML reading for this one fixed shape
/// (`[[fixture]] file = "..." source = "..." sha256 = "..."`), avoiding a
/// full TOML crate dependency for a 3-field repeated table that never
/// needs nested tables or arrays.
fn parse_manifest(text: &str) -> Result<Manifest> {
    let mut fixtures = Vec::new();
    let mut current: Option<(Option<String>, Option<String>, Option<String>)> = None;
    for line in text.lines() {
        let line = line.trim();
        if line.starts_with('#') || line.is_empty() {
            continue;
        }
        if line == "[[fixture]]" {
            if let Some((Some(file), Some(source), Some(sha256))) = current.take() {
                fixtures.push(FixtureEntry { file, source, sha256 });
            }
            current = Some((None, None, None));
            continue;
        }
        if let Some((file, source, sha256)) = &mut current {
            if let Some(v) = parse_kv(line, "file") {
                *file = Some(v);
            } else if let Some(v) = parse_kv(line, "source") {
                *source = Some(v);
            } else if let Some(v) = parse_kv(line, "sha256") {
                *sha256 = Some(v);
            }
        }
    }
    if let Some((Some(file), Some(source), Some(sha256))) = current {
        fixtures.push(FixtureEntry { file, source, sha256 });
    }
    Ok(Manifest { fixture: fixtures })
}

fn parse_kv(line: &str, key: &str) -> Option<String> {
    let prefix = format!("{key} =");
    if !line.starts_with(&prefix) {
        return None;
    }
    let rest = line[prefix.len()..].trim();
    let rest = rest.strip_prefix('"')?;
    let rest = rest.strip_suffix('"')?;
    Some(rest.to_string())
}

fn diff_report(path: &Path) -> Result<()> {
    let parts = office_core::diff::snapshot(path)?;
    let mut names: Vec<&String> = parts.keys().collect();
    names.sort();
    println!("{} parts in {}:", names.len(), path.display());
    for name in names {
        let snap = &parts[name];
        println!("  {name}  ({} bytes, well_formed_xml={})", snap.size, snap.was_well_formed_xml);
    }
    Ok(())
}
