use formula_storage::encryption::is_encrypted_container;
use formula_storage::{CellChange, CellData, CellRange, CellValue, InMemoryKeyProvider, Storage};
use std::sync::Arc;
use tempfile::tempdir;

#[test]
fn encrypted_workbook_round_trip() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("workbook.formula");
    let key_provider = Arc::new(InMemoryKeyProvider::default());

    let storage = Storage::open_encrypted_path(&path, key_provider.clone()).expect("open encrypted");
    let workbook = storage
        .create_workbook("Book1", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet");

    storage
        .apply_cell_changes(&[CellChange {
            sheet_id: sheet.id,
            row: 0,
            col: 0,
            data: CellData {
                value: CellValue::Number(42.0),
                formula: None,
                style: None,
            },
            user_id: None,
        }])
        .expect("apply cells");

    storage.persist().expect("persist encrypted");
    drop(storage);

    let on_disk = std::fs::read(&path).expect("read encrypted file");
    assert!(is_encrypted_container(&on_disk));
    assert!(!on_disk.starts_with(b"SQLite format 3\0"));

    let reopened = Storage::open_encrypted_path(&path, key_provider.clone()).expect("reopen encrypted");
    let sheets = reopened.list_sheets(workbook.id).expect("list sheets");
    assert_eq!(sheets.len(), 1);
    assert_eq!(sheets[0].name, "Sheet1");

    let cells = reopened
        .load_cells_in_range(sheet.id, CellRange::new(0, 5, 0, 5))
        .expect("load cells");
    assert_eq!(cells.len(), 1);
    assert_eq!(cells[0].0, (0, 0));
    assert_eq!(cells[0].1.value, CellValue::Number(42.0));
}

#[test]
fn plaintext_migrates_to_encrypted_on_first_persist() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("workbook_plain.sqlite");

    let plaintext = Storage::open_path(&path).expect("open plaintext");
    let workbook = plaintext
        .create_workbook("Book1", None)
        .expect("create workbook");
    let sheet = plaintext
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet");
    plaintext
        .apply_cell_changes(&[CellChange {
            sheet_id: sheet.id,
            row: 1,
            col: 1,
            data: CellData {
                value: CellValue::String("hello".to_string()),
                formula: None,
                style: None,
            },
            user_id: None,
        }])
        .expect("apply cells");
    drop(plaintext);

    let before = std::fs::read(&path).expect("read plaintext file");
    assert!(before.starts_with(b"SQLite format 3\0"));
    assert!(!is_encrypted_container(&before));

    let key_provider = Arc::new(InMemoryKeyProvider::default());
    let encrypted = Storage::open_encrypted_path(&path, key_provider.clone()).expect("open encrypted");
    encrypted.persist().expect("persist migration");
    drop(encrypted);

    let after = std::fs::read(&path).expect("read migrated file");
    assert!(is_encrypted_container(&after));

    let reopened = Storage::open_encrypted_path(&path, key_provider).expect("reopen migrated");
    let cells = reopened
        .load_cells_in_range(sheet.id, CellRange::new(0, 5, 0, 5))
        .expect("load cells");
    assert_eq!(cells.len(), 1);
    assert_eq!(cells[0].0, (1, 1));
    assert_eq!(cells[0].1.value, CellValue::String("hello".to_string()));
}

