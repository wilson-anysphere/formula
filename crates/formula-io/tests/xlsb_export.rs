use std::path::Path;

use formula_io::{open_workbook, save_workbook};
use formula_model::CellRef;

#[test]
fn xlsb_export_preserves_date_system_1904() {
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../formula-xlsb/tests/fixtures/date1904.xlsb"
    ));
    let wb = open_workbook(fixture_path).expect("open workbook");

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("export.xlsx");
    save_workbook(&wb, &out_path).expect("export workbook");

    let bytes = std::fs::read(&out_path).expect("read exported workbook");
    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load exported workbook");
    assert_eq!(
        doc.workbook.date_system,
        formula_model::DateSystem::Excel1904
    );
}

#[test]
fn xlsb_export_preserves_date_number_format_style() {
    let fixture_path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../formula-xlsb/tests/fixtures_styles/date.xlsb"
    ));
    let wb = open_workbook(fixture_path).expect("open workbook");

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("export.xlsx");
    save_workbook(&wb, &out_path).expect("export workbook");

    let bytes = std::fs::read(&out_path).expect("read exported workbook");
    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load exported workbook");

    let sheet = doc
        .workbook
        .sheet_by_name("Sheet1")
        .unwrap_or_else(|| &doc.workbook.sheets[0]);
    let cell_ref = CellRef::from_a1("A1").expect("valid cell ref");
    let cell = sheet.cell(cell_ref).expect("expected A1 to exist");
    assert_ne!(cell.style_id, 0, "expected A1 to have a non-default style");

    let style = doc
        .workbook
        .styles
        .get(cell.style_id)
        .expect("expected style to exist in workbook style table");
    assert_eq!(style.number_format.as_deref(), Some("m/d/yyyy"));
}

