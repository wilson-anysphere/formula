use desktop::file_io::Workbook;
use desktop::persistence::WorkbookPersistenceLocation;
use desktop::state::AppState;
use serde_json::json;

#[test]
fn add_sheet_creates_persistence_mapping_for_cell_edits() {
    let mut state = AppState::new();
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("Sheet1".to_string());

    state
        .load_workbook_persistent(workbook, WorkbookPersistenceLocation::InMemory)
        .expect("load workbook");

    let sheet = state
        .add_sheet("Sheet".to_string(), None, None, None)
        .expect("add sheet");

    let updates = state
        .set_cell(&sheet.id, 0, 0, Some(json!(123)), None)
        .expect("set cell should succeed on newly-added sheet");

    assert!(
        updates
            .iter()
            .any(|u| u.sheet_id == sheet.id && u.row == 0 && u.col == 0),
        "expected set_cell to report an update for the edited cell"
    );
}

#[test]
fn create_sheet_creates_persistence_mapping_for_cell_edits() {
    let mut state = AppState::new();
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("Sheet1".to_string());

    state
        .load_workbook_persistent(workbook, WorkbookPersistenceLocation::InMemory)
        .expect("load workbook");

    let sheet_id = state
        .create_sheet("Data".to_string())
        .expect("create sheet should succeed");

    let updates = state
        .set_cell(&sheet_id, 0, 0, Some(json!(123)), None)
        .expect("set cell should succeed on newly-created sheet");

    assert!(
        updates
            .iter()
            .any(|u| u.sheet_id == sheet_id && u.row == 0 && u.col == 0),
        "expected set_cell to report an update for the edited cell"
    );
}

#[test]
fn rename_sheet_rewrites_formula_references() {
    let mut state = AppState::new();
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("Sheet1".to_string());
    workbook.add_sheet("Sheet2".to_string());

    state
        .load_workbook_persistent(workbook, WorkbookPersistenceLocation::InMemory)
        .expect("load workbook");

    state
        .set_cell("Sheet2", 0, 0, Some(json!(5)), None)
        .expect("set value");
    state
        .set_cell("Sheet1", 0, 0, None, Some("=Sheet2!A1".to_string()))
        .expect("set formula");

    state
        .rename_sheet("Sheet2", "Data".to_string())
        .expect("rename sheet");

    let cell = state.get_cell("Sheet1", 0, 0).expect("read cell");
    assert_eq!(
        cell.formula.as_deref(),
        Some("=Data!A1"),
        "expected formula to be rewritten after sheet rename"
    );
}

#[test]
fn delete_sheet_rewrites_formula_references_to_ref() {
    let mut state = AppState::new();
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("Sheet1".to_string());
    workbook.add_sheet("Sheet2".to_string());

    state
        .load_workbook_persistent(workbook, WorkbookPersistenceLocation::InMemory)
        .expect("load workbook");

    state
        .set_cell("Sheet1", 0, 0, None, Some("=Sheet2!A1".to_string()))
        .expect("set formula");

    state.delete_sheet("Sheet2").expect("delete sheet");

    let workbook = state.get_workbook().expect("workbook still loaded");
    assert!(
        workbook.sheet("Sheet2").is_none(),
        "expected Sheet2 removed"
    );

    let cell = state.get_cell("Sheet1", 0, 0).expect("read cell");
    assert_eq!(
        cell.formula.as_deref(),
        Some("=#REF!"),
        "expected formula to be rewritten after sheet deletion"
    );
}

#[test]
fn add_sheet_with_id_preserves_stable_id_after_rename_and_delete() {
    let mut state = AppState::new();
    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("Sheet1".to_string());
    workbook.add_sheet("Sheet2".to_string());

    state
        .load_workbook_persistent(workbook, WorkbookPersistenceLocation::InMemory)
        .expect("load workbook");

    // Rename the sheet so the stable id no longer matches the display name.
    state
        .rename_sheet("Sheet2", "Budget".to_string())
        .expect("rename sheet");

    // Delete it, then re-create it with the same stable id.
    state.delete_sheet("Sheet2").expect("delete sheet");
    state
        .add_sheet_with_id("Sheet2".to_string(), "Budget".to_string(), None, None)
        .expect("add sheet with explicit id");

    let info = state.workbook_info().expect("workbook info");
    let sheet = info
        .sheets
        .iter()
        .find(|s| s.id == "Sheet2")
        .expect("expected recreated sheet to exist");

    assert_eq!(sheet.name, "Budget");
    assert_ne!(
        sheet.id, sheet.name,
        "expected explicit id to be preserved (not derived from the sheet name)"
    );

    // Regression: ensure the persistence mapping was created so subsequent cell edits succeed.
    let updates = state
        .set_cell("Sheet2", 0, 0, Some(json!(123)), None)
        .expect("set cell should succeed on recreated sheet");
    assert!(
        updates
            .iter()
            .any(|u| u.sheet_id == "Sheet2" && u.row == 0 && u.col == 0),
        "expected set_cell to report an update for the edited cell"
    );
}
