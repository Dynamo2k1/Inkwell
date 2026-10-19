//! Round-trip tests against a real, Office-authored fixture (python-docx's
//! own `default.docx` template — MIT, vendored from
//! `office-docs-mcp/.venv/.../docx/templates/default.docx`), per the
//! office-rs plan's "real fixtures, not synthetic" discipline.

use std::path::PathBuf;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures")
        .join(name)
}

fn tmp_path(name: &str) -> PathBuf {
    let dir = std::env::temp_dir().join("office-rs-test");
    std::fs::create_dir_all(&dir).unwrap();
    let p = dir.join(name);
    let _ = std::fs::remove_file(&p);
    p
}

#[test]
fn read_real_docx_template() {
    let result = office_docx::read(fixture("docx_default_template.docx")).unwrap();
    // The python-docx default template is intentionally near-empty, but
    // it must open and produce IR sections without erroring.
    assert!(!result.ir.sections.is_empty() || result.plain_text.is_empty());
}

#[test]
fn create_new_docx_fails_if_exists() {
    let dest = tmp_path("create_fail_if_exists.docx");
    std::fs::write(&dest, b"not a real docx, just occupying the path").unwrap();
    let err = office_docx::create(&dest, vec![vec![office_docx::RunSpec::plain("hello")]]);
    assert!(err.is_err(), "create() must refuse to overwrite an existing file");
    std::fs::remove_file(&dest).unwrap();
}

#[test]
fn create_new_docx_and_read_it_back() {
    let dest = tmp_path("created.docx");
    office_docx::create(
        &dest,
        vec![
            vec![office_docx::RunSpec::plain("Hello, world!")],
            vec![office_docx::RunSpec { bold: true, text: "Bold line".into(), ..Default::default() }],
        ],
    )
    .unwrap();

    let read_back = office_docx::read(&dest).unwrap();
    assert!(read_back.plain_text.contains("Hello, world!"));
    assert!(read_back.plain_text.contains("Bold line"));
}

#[test]
fn replace_text_preserves_all_parts() {
    let dest = tmp_path("replace_text_target.docx");
    office_docx::create(
        &dest,
        vec![vec![office_docx::RunSpec::plain("The quick brown fox")]],
    )
    .unwrap();

    // Snapshot the part list ourselves (independent of office-core's own
    // internal snapshot) as an out-of-process-style cross-check per the
    // plan's verification step 3.
    let before_parts = list_zip_parts(&dest);

    let count = office_docx::replace_text(&dest, "quick brown fox", "slow red hen").unwrap();
    assert_eq!(count, 1);

    let after_parts = list_zip_parts(&dest);
    assert_eq!(
        before_parts, after_parts,
        "replace_text must not add or drop any zip part"
    );

    let read_back = office_docx::read(&dest).unwrap();
    assert!(read_back.plain_text.contains("slow red hen"));
    assert!(!read_back.plain_text.contains("quick brown fox"));
}

#[test]
fn append_paragraph_keeps_original_content() {
    let dest = tmp_path("append_target.docx");
    office_docx::create(&dest, vec![vec![office_docx::RunSpec::plain("Original paragraph")]])
        .unwrap();

    let appended = tmp_path("append_result.docx");
    let _ = std::fs::remove_file(&appended);
    office_docx::append_paragraph(
        &dest,
        &appended,
        vec![office_docx::RunSpec::plain("Appended paragraph")],
        None,
    )
    .unwrap();

    let read_back = office_docx::read(&appended).unwrap();
    assert!(read_back.plain_text.contains("Original paragraph"));
    assert!(read_back.plain_text.contains("Appended paragraph"));
}

#[test]
fn insert_table_keeps_original_content() {
    let dest = tmp_path("table_target.docx");
    office_docx::create(&dest, vec![vec![office_docx::RunSpec::plain("Before the table")]])
        .unwrap();

    let result = tmp_path("table_result.docx");
    let _ = std::fs::remove_file(&result);
    office_docx::insert_table(
        &dest,
        &result,
        vec![
            vec!["Header A".to_string(), "Header B".to_string()],
            vec!["1".to_string(), "2".to_string()],
        ],
    )
    .unwrap();

    let read_back = office_docx::read(&result).unwrap();
    assert!(read_back.plain_text.contains("Before the table"));
    assert!(read_back.plain_text.contains("Header A"));
}

#[test]
fn insert_image_appends_real_image_part() {
    let dest = tmp_path("image_target.docx");
    office_docx::create(&dest, vec![vec![office_docx::RunSpec::plain("Before the image")]]).unwrap();

    let png_bytes = std::fs::read(fixture("tiny_red.png")).unwrap();
    let result = tmp_path("image_result.docx");
    let _ = std::fs::remove_file(&result);
    office_docx::insert_image(
        &dest,
        &result,
        png_bytes,
        office_oxide::ir::ImageFormat::Png,
        Some("a tiny red square".to_string()),
    )
    .unwrap();

    let read_back = office_docx::read(&result).unwrap();
    assert!(read_back.plain_text.contains("Before the image"));
    // A real media part must exist in the saved zip, not just an IR entry.
    let parts = list_zip_parts(&result);
    assert!(parts.iter().any(|p| p.starts_with("word/media/")), "expected an embedded media part, got: {parts:?}");
}

fn list_zip_parts(path: &std::path::Path) -> Vec<String> {
    let file = std::fs::File::open(path).unwrap();
    let mut zip = zip::ZipArchive::new(file).unwrap();
    let mut names: Vec<String> = (0..zip.len()).map(|i| zip.by_index(i).unwrap().name().to_string()).collect();
    names.sort();
    names
}
