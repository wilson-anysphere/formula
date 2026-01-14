use desktop::file_io::read_xlsx_blocking;
use desktop::persistence::WorkbookPersistenceLocation;
use desktop::state::AppState;
use std::path::PathBuf;

#[test]
fn xlsx_import_populates_sheet_col_properties_width_and_hidden() {
    let fixture = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../../fixtures/xlsx/basic/row-col-attrs.xlsx");
    let workbook = read_xlsx_blocking(&fixture).expect("read fixture workbook");

    let mut state = AppState::new();
    state
        .load_workbook_persistent(workbook, WorkbookPersistenceLocation::InMemory)
        .expect("load workbook into persistence");

    let sheet_uuid = state.persistent_sheet_uuid("Sheet1").expect("sheet uuid");
    let storage = state.persistent_storage().expect("persistent storage");
    let sheet = storage
        .get_sheet_model_worksheet(sheet_uuid)
        .expect("read sheet model json")
        .expect("sheet model worksheet should be present");

    // `fixtures/xlsx/basic/row-col-attrs.xlsx` defines:
    // - `<col min=\"2\" max=\"2\" width=\"25\" customWidth=\"1\"/>` (B column, 0-based col 1)
    // - `<col min=\"3\" max=\"3\" hidden=\"1\"/>` (C column, 0-based col 2)
    let col_b = sheet
        .col_properties
        .get(&1)
        .expect("expected column B properties in model");
    assert!(
        col_b.width.is_some_and(|w| (w - 25.0).abs() <= 1e-6),
        "expected column B width=25, got {:?}",
        col_b.width
    );
    assert!(
        !col_b.hidden,
        "expected column B not hidden; got hidden=true"
    );

    let col_c = sheet
        .col_properties
        .get(&2)
        .expect("expected column C properties in model");
    assert!(col_c.hidden, "expected column C hidden=true");
}

