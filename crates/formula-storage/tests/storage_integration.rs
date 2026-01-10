use formula_storage::{AutoSaveConfig, AutoSaveManager, CellChange, CellData, CellRange, CellValue, Storage, Style};
use serde_json::json;
use std::time::Duration;

#[test]
fn save_load_round_trip_shared_memory() {
    // Use a shared in-memory database so we can open a second connection and
    // simulate reloading the workbook from disk.
    let uri = "file:round_trip?mode=memory&cache=shared";

    let storage1 = Storage::open_uri(uri).expect("open storage");
    let workbook = storage1
        .create_workbook("Book1", None)
        .expect("create workbook");
    let sheet = storage1
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet");

    storage1
        .apply_cell_changes(&[
            CellChange {
                sheet_id: sheet.id,
                row: 0,
                col: 0,
                data: CellData {
                    value: CellValue::Number(42.0),
                    formula: None,
                    style: None,
                },
                user_id: Some("test-user".to_string()),
            },
            CellChange {
                sheet_id: sheet.id,
                row: 1,
                col: 1,
                data: CellData {
                    value: CellValue::Text("hello".to_string()),
                    formula: None,
                    style: None,
                },
                user_id: Some("test-user".to_string()),
            },
        ])
        .expect("persist cells");

    // Open a second handle to the same shared memory DB.
    let storage2 = Storage::open_uri(uri).expect("open second storage");
    let sheets = storage2.list_sheets(workbook.id).expect("list sheets");
    assert_eq!(sheets.len(), 1);
    assert_eq!(sheets[0].name, "Sheet1");

    let cells = storage2
        .load_cells_in_range(sheet.id, CellRange::new(0, 10, 0, 10))
        .expect("load cells");

    assert_eq!(cells.len(), 2);
    assert_eq!(cells[0].0, (0, 0));
    assert_eq!(cells[0].1.value, CellValue::Number(42.0));
    assert_eq!(cells[1].0, (1, 1));
    assert_eq!(cells[1].1.value, CellValue::Text("hello".to_string()));

    // Keep storage1 alive for the lifetime of the test to ensure the shared
    // in-memory DB isn't dropped.
    std::mem::drop(storage1);
}

#[test]
fn sparse_storage_only_persists_non_empty_cells() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage.create_workbook("Book", None).expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet", 0, None)
        .expect("create sheet");

    storage
        .apply_cell_changes(&[
            CellChange {
                sheet_id: sheet.id,
                row: 0,
                col: 0,
                data: CellData {
                    value: CellValue::Number(1.0),
                    formula: None,
                    style: None,
                },
                user_id: None,
            },
            CellChange {
                sheet_id: sheet.id,
                row: 999_999,
                col: 999,
                data: CellData {
                    value: CellValue::Number(2.0),
                    formula: None,
                    style: None,
                },
                user_id: None,
            },
        ])
        .expect("persist cells");

    assert_eq!(storage.cell_count(sheet.id).unwrap(), 2);

    // Deleting a cell should remove its row from the `cells` table.
    storage
        .apply_cell_changes(&[CellChange {
            sheet_id: sheet.id,
            row: 0,
            col: 0,
            data: CellData::empty(),
            user_id: None,
        }])
        .expect("delete cell");

    assert_eq!(storage.cell_count(sheet.id).unwrap(), 1);
}

#[test]
fn styles_are_deduplicated() {
    let storage = Storage::open_in_memory().expect("open storage");

    let style_a = Style {
        font_id: None,
        fill_id: None,
        border_id: None,
        number_format: Some("0.0".to_string()),
        alignment: Some(json!({"b": 2, "a": 1})),
        protection: None,
    };
    let style_b = Style {
        font_id: None,
        fill_id: None,
        border_id: None,
        number_format: Some("0.0".to_string()),
        alignment: Some(json!({"a": 1, "b": 2})),
        protection: None,
    };

    let id_a = storage.get_or_insert_style(&style_a).expect("insert style");
    let id_b = storage.get_or_insert_style(&style_b).expect("dedup style");
    assert_eq!(id_a, id_b);

    let style_c = Style {
        number_format: Some("0.00".to_string()),
        ..style_a
    };
    let id_c = storage.get_or_insert_style(&style_c).expect("insert other style");
    assert_ne!(id_a, id_c);
}

#[test]
fn change_log_records_cell_operations() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage.create_workbook("Book", None).expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet", 0, None)
        .expect("create sheet");

    storage
        .apply_cell_changes(&[CellChange {
            sheet_id: sheet.id,
            row: 0,
            col: 0,
            data: CellData {
                value: CellValue::Text("hello".to_string()),
                formula: None,
                style: None,
            },
            user_id: Some("alice".to_string()),
        }])
        .expect("set cell");

    assert_eq!(storage.change_log_count(sheet.id).unwrap(), 1);
    let latest = storage.latest_change(sheet.id).unwrap().expect("latest");
    assert_eq!(latest.operation, "set_cell");
    assert_eq!(latest.user_id.as_deref(), Some("alice"));
    assert_eq!(latest.target, json!({"row": 0, "col": 0}));

    storage
        .apply_cell_changes(&[CellChange {
            sheet_id: sheet.id,
            row: 0,
            col: 0,
            data: CellData::empty(),
            user_id: Some("alice".to_string()),
        }])
        .expect("delete cell");

    assert_eq!(storage.change_log_count(sheet.id).unwrap(), 2);
    let latest = storage.latest_change(sheet.id).unwrap().expect("latest");
    assert_eq!(latest.operation, "delete_cell");
}

#[tokio::test(flavor = "current_thread")]
async fn autosave_batches_changes() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage.create_workbook("Book", None).expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet", 0, None)
        .expect("create sheet");

    let autosave = AutoSaveManager::spawn(
        storage.clone(),
        AutoSaveConfig {
            save_delay: Duration::from_millis(50),
            max_delay: Duration::from_millis(200),
        },
    );

    autosave.record_change(CellChange {
        sheet_id: sheet.id,
        row: 0,
        col: 0,
        data: CellData {
            value: CellValue::Number(10.0),
            formula: None,
            style: None,
        },
        user_id: None,
    });

    autosave.record_change(CellChange {
        sheet_id: sheet.id,
        row: 0,
        col: 1,
        data: CellData {
            value: CellValue::Number(20.0),
            formula: None,
            style: None,
        },
        user_id: None,
    });

    // Wait long enough for the debounce timer to fire.
    tokio::time::sleep(Duration::from_millis(120)).await;

    // Ensure pending changes are flushed and then validate we only persisted once.
    autosave.flush().await.expect("flush");
    assert_eq!(autosave.save_count(), 1);

    let cells = storage
        .load_cells_in_range(sheet.id, CellRange::new(0, 0, 0, 10))
        .expect("load cells");
    assert_eq!(cells.len(), 2);

    autosave.shutdown().await.expect("shutdown");
}
