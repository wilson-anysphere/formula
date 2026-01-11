use formula_desktop_tauri::file_io::{read_xlsx_blocking, Workbook};
use formula_desktop_tauri::persistence::{write_xlsx_from_storage, WorkbookPersistenceLocation};
use formula_desktop_tauri::state::{AppState, CellScalar};
use formula_storage::{CellRange, CellValue, Storage};
use serde_json::json;
use std::path::PathBuf;

#[tokio::test]
async fn autosave_persists_and_exports_round_trip() {
    let tmp_dir = tempfile::tempdir().expect("temp dir");
    let db_path = tmp_dir.path().join("autosave.sqlite");

    let mut workbook = Workbook::new_empty(None);
    workbook.add_sheet("Sheet1".to_string());
    let sheet_id = workbook.sheets[0].id.clone();

    let mut state = AppState::new();
    state
        .load_workbook_persistent(workbook, WorkbookPersistenceLocation::OnDisk(db_path.clone()))
        .expect("load persistent workbook");

    state
        .set_cell(&sheet_id, 0, 0, Some(json!(42.0)), None)
        .expect("set A1");
    state
        .set_cell(&sheet_id, 0, 1, None, Some("=A1+1".to_string()))
        .expect("set B1 formula");
    assert_eq!(
        state.get_cell(&sheet_id, 0, 1).expect("get B1").value,
        CellScalar::Number(43.0)
    );

    let autosave = state.autosave_manager().expect("autosave manager");
    autosave.flush().await.expect("flush autosave");

    let storage = Storage::open_path(&db_path).expect("open storage");
    let workbooks = storage.list_workbooks().expect("list workbooks");
    assert_eq!(workbooks.len(), 1);
    let workbook_id = workbooks[0].id;

    let sheets = storage.list_sheets(workbook_id).expect("list sheets");
    let sheet_uuid = sheets
        .iter()
        .find(|s| s.name == "Sheet1")
        .expect("sheet1 meta")
        .id;

    let persisted = storage
        .load_cells_in_range(sheet_uuid, CellRange::new(0, 0, 0, 1))
        .expect("load persisted cells");
    assert_eq!(persisted.len(), 2);
    assert!(
        persisted
            .iter()
            .any(|((r, c), snap)| *r == 0 && *c == 0 && snap.value == CellValue::Number(42.0)),
        "expected A1=42 in persisted cells, got {persisted:?}"
    );
    assert!(
        persisted.iter().any(|((r, c), snap)| {
            *r == 0 && *c == 1 && snap.formula.as_deref() == Some("A1+1")
        }),
        "expected B1 formula in persisted cells, got {persisted:?}"
    );

    // Clearing a cell should remove it from SQLite.
    state
        .set_cell(&sheet_id, 0, 0, None, None)
        .expect("clear A1");
    autosave.flush().await.expect("flush autosave after clear");
    assert_eq!(storage.cell_count(sheet_uuid).expect("cell count"), 1);

    // Recovery: load a new AppState from the same autosave DB.
    let mut template = Workbook::new_empty(None);
    template.add_sheet("Sheet1".to_string());
    let mut recovered = AppState::new();
    recovered
        .load_workbook_persistent(template, WorkbookPersistenceLocation::OnDisk(db_path.clone()))
        .expect("load recovered workbook");

    assert_eq!(
        recovered.get_cell(&sheet_id, 0, 0).expect("get A1").value,
        CellScalar::Empty
    );
    assert_eq!(
        recovered.get_cell(&sheet_id, 0, 1).expect("get B1").value,
        CellScalar::Number(1.0),
        "empty cells are treated as 0 in arithmetic"
    );

    // Export through the same storage→xlsx path as the desktop `save_workbook` command.
    let xlsx_path = tmp_dir.path().join("export.xlsx");
    let export_storage = recovered.persistent_storage().expect("storage handle");
    let export_id = recovered.persistent_workbook_id().expect("workbook id");
    let export_meta = recovered.get_workbook().expect("workbook").clone();
    recovered
        .autosave_manager()
        .expect("autosave")
        .flush()
        .await
        .expect("flush before export");

    write_xlsx_from_storage(&export_storage, export_id, &export_meta, &xlsx_path).expect("export xlsx");

    let reopened = read_xlsx_blocking(&xlsx_path).expect("read exported xlsx");
    let reopened_sheet_id = reopened.sheets[0].id.clone();
    let mut reopened_state = AppState::new();
    reopened_state.load_workbook(reopened);

    let cell = reopened_state
        .get_cell(&reopened_sheet_id, 0, 1)
        .expect("get B1 after reopen");
    assert_eq!(cell.formula.as_deref(), Some("=A1+1"));
    assert_eq!(cell.value, CellScalar::Number(1.0));
}

#[test]
fn xlsx_import_uses_origin_bytes_to_preserve_styles_in_storage() {
    let fixture = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../../fixtures/xlsx/styles/styles.xlsx");
    let workbook = read_xlsx_blocking(&fixture).expect("read fixture");
    assert!(
        workbook.origin_xlsx_bytes.is_some(),
        "fixture should load with origin_xlsx_bytes"
    );

    let tmp_dir = tempfile::tempdir().expect("temp dir");
    let db_path = tmp_dir.path().join("autosave.sqlite");

    let mut state = AppState::new();
    state
        .load_workbook_persistent(workbook, WorkbookPersistenceLocation::OnDisk(db_path))
        .expect("load persistent workbook");

    let storage = state.persistent_storage().expect("storage");
    let workbook_id = state.persistent_workbook_id().expect("workbook id");
    let model = storage
        .export_model_workbook(workbook_id)
        .expect("export model");

    assert!(
        model.styles.len() > 1,
        "expected more than the default style in imported workbook"
    );
}
