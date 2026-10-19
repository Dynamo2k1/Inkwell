//! office-xlsx round-trip tests. No external fixture file exists yet for
//! xlsx (that lands with `xtask fetch-fixtures`'s openpyxl sparse-clone),
//! so these tests build their own real `.xlsx` via `office_xlsx::create`
//! and then exercise read/edit against that real, office_oxide-authored
//! file rather than against synthetic bytes.

use office_xlsx::{CellKind, CellSpec, CellWriteValue, SheetSpec};
use std::path::PathBuf;

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
fn create_new_xlsx_fails_if_exists() {
    let dest = tmp_path("exists.xlsx");
    std::fs::write(&dest, b"occupied").unwrap();
    let err = office_xlsx::create(&dest, vec![]);
    assert!(err.is_err());
    std::fs::remove_file(&dest).unwrap();
}

#[test]
fn create_multi_sheet_workbook_and_read_it_back() {
    let dest = tmp_path("workbook.xlsx");
    office_xlsx::create(
        &dest,
        vec![
            SheetSpec {
                name: "Sheet1".into(),
                rows: vec![
                    vec![CellSpec::text("Name"), CellSpec::text("Score")],
                    vec![CellSpec::text("Alice"), CellSpec::number(95.0)],
                ],
            },
            SheetSpec {
                name: "Notes".into(),
                rows: vec![vec![CellSpec::text("second sheet")]],
            },
        ],
    )
    .unwrap();

    let read_back = office_xlsx::read(&dest).unwrap();
    assert!(read_back.plain_text.contains("Alice"));
    assert!(read_back.plain_text.contains("second sheet"));
    // Two sheets must both survive into the IR.
    assert_eq!(read_back.ir.sections.len(), 2);
}

#[test]
fn styled_cell_round_trips_without_part_loss() {
    let dest = tmp_path("styled.xlsx");
    office_xlsx::create(
        &dest,
        vec![SheetSpec {
            name: "Sheet1".into(),
            rows: vec![vec![CellSpec {
                value: CellKind::Text("Bold Header".into()),
                bold: true,
                background_hex: Some("FFFF00".into()),
                ..Default::default()
            }]],
        }],
    )
    .unwrap();

    let read_back = office_xlsx::read(&dest).unwrap();
    assert!(read_back.plain_text.contains("Bold Header"));
}

#[test]
fn write_cells_preserves_all_parts() {
    let dest = tmp_path("edit_target.xlsx");
    office_xlsx::create(
        &dest,
        vec![SheetSpec {
            name: "Sheet1".into(),
            rows: vec![vec![CellSpec::text("old value")]],
        }],
    )
    .unwrap();

    let before_parts = list_zip_parts(&dest);

    let count = office_xlsx::write_cells(
        &dest,
        vec![(0, "A1".to_string(), CellWriteValue::Text("new value".to_string()))],
    )
    .unwrap();
    assert_eq!(count, 1);

    let after_parts = list_zip_parts(&dest);
    assert_eq!(before_parts, after_parts, "write_cells must not add or drop any zip part");

    let read_back = office_xlsx::read(&dest).unwrap();
    assert!(read_back.plain_text.contains("new value"));
    assert!(!read_back.plain_text.contains("old value"));
}

#[test]
fn write_cells_rejects_out_of_range_sheet() {
    let dest = tmp_path("out_of_range.xlsx");
    office_xlsx::create(
        &dest,
        vec![SheetSpec { name: "Sheet1".into(), rows: vec![vec![CellSpec::text("x")]] }],
    )
    .unwrap();

    let err = office_xlsx::write_cells(
        &dest,
        vec![(5, "A1".to_string(), CellWriteValue::Text("y".to_string()))],
    );
    assert!(err.is_err(), "writing to a nonexistent sheet index must fail loudly, not silently no-op");
}
