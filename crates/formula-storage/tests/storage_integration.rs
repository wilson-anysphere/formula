use formula_storage::{AutoSaveConfig, AutoSaveManager, CellChange, CellData, CellRange, CellValue, Storage};
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

