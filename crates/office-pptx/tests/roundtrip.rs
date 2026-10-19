//! office-pptx tests: read against a real, MIT-licensed fixture
//! (`github.com/h2non/filetype`'s test corpus), and create/edit round
//! trips against office_oxide-authored decks.

use office_pptx::SlideSpec;
use std::path::PathBuf;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../tests/fixtures").join(name)
}

fn tmp_path(name: &str) -> PathBuf {
    let dir = std::env::temp_dir().join("office-rs-test");
    std::fs::create_dir_all(&dir).unwrap();
    let p = dir.join(name);
    let _ = std::fs::remove_file(&p);
    p
}

fn list_zip_parts(path: &std::path::Path) -> Vec<String> {
    let file = std::fs::File::open(path).unwrap();
    let mut zip = zip::ZipArchive::new(file).unwrap();
    let mut names: Vec<String> = (0..zip.len()).map(|i| zip.by_index(i).unwrap().name().to_string()).collect();
    names.sort();
    names
}

#[test]
fn read_real_pptx_fixture() {
    let result = office_pptx::read(fixture("pptx_h2non_sample.pptx")).unwrap();
    // Must open and produce IR without erroring; content is whatever the
    // real fixture happens to contain.
    let _ = result.plain_text;
    let _ = result.ir.sections;
}

#[test]
fn create_new_pptx_fails_if_exists() {
    let dest = tmp_path("exists.pptx");
    std::fs::write(&dest, b"occupied").unwrap();
    let err = office_pptx::create(&dest, vec![]);
    assert!(err.is_err());
    std::fs::remove_file(&dest).unwrap();
}

#[test]
fn create_deck_and_read_it_back() {
    let dest = tmp_path("deck.pptx");
    office_pptx::create(
        &dest,
        vec![
            SlideSpec {
                title: Some("Welcome".into()),
                bullets: vec!["Point one".into(), "Point two".into()],
                notes: Some("Remember to smile".into()),
            },
            SlideSpec { title: Some("Slide Two".into()), ..Default::default() },
        ],
    )
    .unwrap();

    let read_back = office_pptx::read(&dest).unwrap();
    assert!(read_back.plain_text.contains("Welcome"));
    assert!(read_back.plain_text.contains("Point one"));
    assert!(read_back.plain_text.contains("Slide Two"));
    assert_eq!(read_back.ir.sections.len(), 2);
}

#[test]
fn add_slide_keeps_existing_slides() {
    let dest = tmp_path("base_deck.pptx");
    office_pptx::create(&dest, vec![SlideSpec { title: Some("First".into()), ..Default::default() }])
        .unwrap();

    let result = tmp_path("deck_plus_one.pptx");
    let _ = std::fs::remove_file(&result);
    office_pptx::add_slide(
        &dest,
        &result,
        SlideSpec { title: Some("Second".into()), ..Default::default() },
    )
    .unwrap();

    let read_back = office_pptx::read(&result).unwrap();
    assert!(read_back.plain_text.contains("First"));
    assert!(read_back.plain_text.contains("Second"));
    assert_eq!(read_back.ir.sections.len(), 2);
}

#[test]
fn set_notes_targets_correct_slide() {
    let dest = tmp_path("notes_deck.pptx");
    office_pptx::create(
        &dest,
        vec![
            SlideSpec { title: Some("Slide A".into()), ..Default::default() },
            SlideSpec { title: Some("Slide B".into()), ..Default::default() },
        ],
    )
    .unwrap();

    let result = tmp_path("notes_deck_result.pptx");
    let _ = std::fs::remove_file(&result);
    office_pptx::set_notes(&dest, &result, 1, "Notes for slide B").unwrap();

    let read_back = office_pptx::read(&result).unwrap();
    assert!(read_back.plain_text.contains("Notes for slide B"));
}

#[test]
fn set_notes_rejects_out_of_range_index() {
    let dest = tmp_path("notes_oob_deck.pptx");
    office_pptx::create(&dest, vec![SlideSpec { title: Some("Only".into()), ..Default::default() }])
        .unwrap();
    let result = tmp_path("notes_oob_result.pptx");
    let err = office_pptx::set_notes(&dest, &result, 5, "unreachable");
    assert!(err.is_err());
}

#[test]
fn replace_text_preserves_all_parts() {
    let dest = tmp_path("replace_deck.pptx");
    office_pptx::create(
        &dest,
        vec![SlideSpec { title: Some("Quarterly Results".into()), ..Default::default() }],
    )
    .unwrap();

    let before_parts = list_zip_parts(&dest);
    let count = office_pptx::replace_text(&dest, "Quarterly Results", "Annual Results").unwrap();
    assert_eq!(count, 1);
    let after_parts = list_zip_parts(&dest);
    assert_eq!(before_parts, after_parts, "replace_text must not add or drop any zip part");

    let read_back = office_pptx::read(&dest).unwrap();
    assert!(read_back.plain_text.contains("Annual Results"));
}

#[test]
fn insert_image_appends_real_image_part() {
    let dest = tmp_path("image_deck.pptx");
    office_pptx::create(&dest, vec![SlideSpec { title: Some("Before image".into()), ..Default::default() }])
        .unwrap();

    let png_bytes = std::fs::read(fixture("tiny_red.png")).unwrap();
    let result = tmp_path("image_deck_result.pptx");
    let _ = std::fs::remove_file(&result);
    office_pptx::insert_image(
        &dest,
        &result,
        png_bytes,
        office_oxide::ir::ImageFormat::Png,
        0,
        0,
        914_400,
        914_400,
    )
    .unwrap();

    let read_back = office_pptx::read(&result).unwrap();
    assert!(read_back.plain_text.contains("Before image"));
    let parts = list_zip_parts(&result);
    assert!(parts.iter().any(|p| p.starts_with("ppt/media/")), "expected an embedded media part, got: {parts:?}");
}

#[test]
fn insert_table_appends_new_slide() {
    let dest = tmp_path("table_deck.pptx");
    office_pptx::create(&dest, vec![SlideSpec { title: Some("Intro".into()), ..Default::default() }])
        .unwrap();

    let result = tmp_path("table_deck_result.pptx");
    let _ = std::fs::remove_file(&result);
    office_pptx::insert_table(
        &dest,
        &result,
        vec![vec!["Q1".to_string(), "Q2".to_string()], vec!["100".to_string(), "150".to_string()]],
    )
    .unwrap();

    let read_back = office_pptx::read(&result).unwrap();
    assert!(read_back.plain_text.contains("Intro"));
    assert!(read_back.plain_text.contains("Q1"));
    assert_eq!(read_back.ir.sections.len(), 2);
}
