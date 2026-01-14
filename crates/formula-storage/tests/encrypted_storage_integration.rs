use formula_model::{
    Cell as ModelCell, CellRef, CellValue as ModelCellValue, Workbook as ModelWorkbook,
};
use formula_storage::encryption::is_encrypted_container;
use formula_storage::{storage::StorageError, AutoSaveConfig, AutoSaveManager, EncryptionError};
use formula_storage::{
    CellChange, CellData, CellRange, CellValue, ImportModelWorkbookOptions, InMemoryKeyProvider,
    KeyProvider, MemoryManager, MemoryManagerConfig, Storage,
};
use std::sync::Arc;
use std::time::Duration;
use tempfile::tempdir;

#[test]
fn encrypted_workbook_round_trip() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("workbook.formula");
    let key_provider = Arc::new(InMemoryKeyProvider::default());

    let storage =
        Storage::open_encrypted_path(&path, key_provider.clone()).expect("open encrypted");
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

    let reopened =
        Storage::open_encrypted_path(&path, key_provider.clone()).expect("reopen encrypted");
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
    let encrypted =
        Storage::open_encrypted_path(&path, key_provider.clone()).expect("open encrypted");
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

#[test]
fn encrypted_persist_creates_parent_directories() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("nested/dir/workbook.formula");
    let key_provider = Arc::new(InMemoryKeyProvider::default());

    let storage = Storage::open_encrypted_path(&path, key_provider).expect("open encrypted");
    storage
        .create_workbook("Book1", None)
        .expect("create workbook");
    storage.persist().expect("persist should create dirs");
    assert!(path.exists(), "expected encrypted workbook file to exist");
}

#[test]
fn encrypted_open_with_wrong_key_fails() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("workbook.formula");

    let good_provider = Arc::new(InMemoryKeyProvider::default());
    let storage = Storage::open_encrypted_path(&path, good_provider).expect("open encrypted");
    storage
        .create_workbook("Book1", None)
        .expect("create workbook");
    storage.persist().expect("persist encrypted");
    drop(storage);

    // Use a different key (but same key version) to ensure decrypt fails safely.
    let wrong_keyring = formula_storage::KeyRing::from_key(1, [0xAA; 32]);
    let wrong_provider = Arc::new(InMemoryKeyProvider::new(Some(wrong_keyring)));

    let err = Storage::open_encrypted_path(&path, wrong_provider)
        .expect_err("open should fail with wrong key");
    match err {
        StorageError::Encryption(EncryptionError::Aead) => {}
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn encrypted_persist_removes_sqlite_sidecars() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("workbook.formula");
    let key_provider = Arc::new(InMemoryKeyProvider::default());

    // Simulate leftover SQLite sidecar files that could contain plaintext.
    for suffix in ["-wal", "-shm", "-journal"] {
        let sidecar = format!("{}{}", path.display(), suffix);
        std::fs::write(sidecar, b"plaintext").expect("write sidecar");
    }

    let storage = Storage::open_encrypted_path(&path, key_provider).expect("open encrypted");
    storage
        .create_workbook("Book1", None)
        .expect("create workbook");
    storage.persist().expect("persist");

    for suffix in ["-wal", "-shm", "-journal"] {
        let sidecar = format!("{}{}", path.display(), suffix);
        assert!(
            !std::path::Path::new(&sidecar).exists(),
            "expected sidecar {sidecar} to be removed"
        );
    }
}

#[tokio::test(flavor = "current_thread")]
async fn encrypted_autosave_persists_to_disk() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("workbook.formula");
    let key_provider = Arc::new(InMemoryKeyProvider::default());

    let storage =
        Storage::open_encrypted_path(&path, key_provider.clone()).expect("open encrypted");
    let workbook = storage
        .create_workbook("Book1", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet");

    let memory = MemoryManager::new(
        storage.clone(),
        MemoryManagerConfig {
            max_memory_bytes: 128 * 1024,
            max_pages: 128,
            eviction_watermark: 0.8,
            rows_per_page: 64,
            cols_per_page: 64,
        },
    );

    let autosave = AutoSaveManager::spawn(
        memory,
        AutoSaveConfig {
            save_delay: Duration::from_millis(50),
            max_delay: Duration::from_millis(200),
        },
    );

    autosave
        .record_change(CellChange {
            sheet_id: sheet.id,
            row: 0,
            col: 0,
            data: CellData {
                value: CellValue::Number(42.0),
                formula: None,
                style: None,
            },
            user_id: None,
        })
        .expect("record change");

    tokio::time::sleep(Duration::from_millis(120)).await;
    autosave.flush().await.expect("flush");
    autosave.shutdown().await.expect("shutdown");

    let reopened = Storage::open_encrypted_path(&path, key_provider).expect("reopen encrypted");
    let cells = reopened
        .load_cells_in_range(sheet.id, CellRange::new(0, 5, 0, 5))
        .expect("load cells");
    assert_eq!(cells.len(), 1);
    assert_eq!(cells[0].0, (0, 0));
    assert_eq!(cells[0].1.value, CellValue::Number(42.0));
}

#[test]
fn encrypted_workbook_survives_key_rotation() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("workbook.formula");
    let key_provider = Arc::new(InMemoryKeyProvider::default());

    let storage =
        Storage::open_encrypted_path(&path, key_provider.clone()).expect("open encrypted");
    let workbook = storage
        .create_workbook("Book1", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet");
    storage
        .apply_cell_changes(&[CellChange {
            sheet_id: sheet.id,
            row: 2,
            col: 3,
            data: CellData {
                value: CellValue::Number(123.0),
                formula: None,
                style: None,
            },
            user_id: None,
        }])
        .expect("apply cells");
    storage.persist().expect("persist v1");
    drop(storage);

    let bytes_v1 = std::fs::read(&path).expect("read v1");
    assert_eq!(&bytes_v1[..8], b"FMLENC01");
    let key_version_v1 = u32::from_be_bytes(bytes_v1[8..12].try_into().expect("key version bytes"));
    assert_eq!(key_version_v1, 1);

    // Rotate keys in the provider while retaining old versions.
    let mut ring = key_provider
        .keyring()
        .expect("keyring should exist after persist");
    ring.rotate();
    key_provider
        .store_keyring(&ring)
        .expect("store rotated keyring");

    // Reopen with rotated keyring; should still decrypt v1.
    let reopened =
        Storage::open_encrypted_path(&path, key_provider.clone()).expect("reopen encrypted");
    let cells = reopened
        .load_cells_in_range(sheet.id, CellRange::new(0, 10, 0, 10))
        .expect("load cells");
    assert_eq!(cells.len(), 1);
    assert_eq!(cells[0].0, (2, 3));
    assert_eq!(cells[0].1.value, CellValue::Number(123.0));

    // Persist again; should now re-encrypt with key version 2.
    reopened.persist().expect("persist v2");
    drop(reopened);

    let bytes_v2 = std::fs::read(&path).expect("read v2");
    assert_eq!(&bytes_v2[..8], b"FMLENC01");
    let key_version_v2 = u32::from_be_bytes(bytes_v2[8..12].try_into().expect("key version bytes"));
    assert_eq!(key_version_v2, 2);

    let reopened_again =
        Storage::open_encrypted_path(&path, key_provider).expect("reopen after v2 persist");
    let cells = reopened_again
        .load_cells_in_range(sheet.id, CellRange::new(0, 10, 0, 10))
        .expect("load cells");
    assert_eq!(cells.len(), 1);
    assert_eq!(cells[0].0, (2, 3));
    assert_eq!(cells[0].1.value, CellValue::Number(123.0));
}

#[test]
fn rotate_encryption_key_api_reencrypts_file() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("workbook.formula");
    let key_provider = Arc::new(InMemoryKeyProvider::default());

    let storage =
        Storage::open_encrypted_path(&path, key_provider.clone()).expect("open encrypted");
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
                value: CellValue::Number(99.0),
                formula: None,
                style: None,
            },
            user_id: None,
        }])
        .expect("apply cells");
    storage.persist().expect("persist v1");

    let bytes_v1 = std::fs::read(&path).expect("read v1");
    let key_version_v1 = u32::from_be_bytes(bytes_v1[8..12].try_into().expect("key version bytes"));
    assert_eq!(key_version_v1, 1);

    let rotated = storage
        .rotate_encryption_key()
        .expect("rotate encryption key")
        .expect("should be encrypted");
    assert_eq!(rotated, 2);

    let bytes_v2 = std::fs::read(&path).expect("read v2");
    let key_version_v2 = u32::from_be_bytes(bytes_v2[8..12].try_into().expect("key version bytes"));
    assert_eq!(key_version_v2, 2);

    let reopened = Storage::open_encrypted_path(&path, Arc::new(InMemoryKeyProvider::default()))
        .expect_err("different provider should fail due to missing keyring");
    assert!(matches!(reopened, StorageError::Encryption(_)));

    drop(storage);

    // Ensure the original provider can still decrypt and data survives the re-encryption.
    let reopened = Storage::open_encrypted_path(&path, key_provider).expect("reopen after rotate");
    let cells = reopened
        .load_cells_in_range(sheet.id, CellRange::new(0, 5, 0, 5))
        .expect("load cells");
    assert_eq!(cells.len(), 1);
    assert_eq!(cells[0].0, (0, 0));
    assert_eq!(cells[0].1.value, CellValue::Number(99.0));
}

#[test]
fn encrypted_model_workbook_round_trip() {
    let dir = tempdir().expect("tempdir");
    let path = dir.path().join("model.formula");
    let key_provider = Arc::new(InMemoryKeyProvider::default());

    let mut model = ModelWorkbook::new();
    model.id = 123;
    model.schema_version = formula_model::SCHEMA_VERSION;
    let sheet_id = model.add_sheet("Sheet1").expect("add sheet");
    {
        let sheet = model.sheet_mut(sheet_id).expect("sheet");
        sheet.set_cell(
            CellRef::new(0, 0),
            ModelCell {
                value: ModelCellValue::Number(7.0),
                formula: None,
                phonetic: None,
                style_id: 0,
                phonetic: None,
            },
        );
        sheet.set_cell(
            CellRef::new(0, 1),
            ModelCell {
                value: ModelCellValue::Empty,
                formula: Some("SUM(A1)".to_string()),
                phonetic: None,
                style_id: 0,
                phonetic: None,
            },
        );
    }

    let storage =
        Storage::open_encrypted_path(&path, key_provider.clone()).expect("open encrypted");
    let meta = storage
        .import_model_workbook(&model, ImportModelWorkbookOptions::new("Book"))
        .expect("import model workbook");
    storage.persist().expect("persist encrypted");
    drop(storage);

    let reopened = Storage::open_encrypted_path(&path, key_provider).expect("reopen encrypted");
    let exported = reopened
        .export_model_workbook(meta.id)
        .expect("export model workbook");

    assert_eq!(exported.id, model.id);
    assert_eq!(exported.schema_version, model.schema_version);
    assert_eq!(exported.sheets.len(), 1);
    let sheet = exported.sheet_by_name("Sheet1").expect("sheet exists");
    let cell_a1 = sheet.cell(CellRef::new(0, 0)).expect("A1");
    assert_eq!(cell_a1.value, ModelCellValue::Number(7.0));
    let cell_b1 = sheet.cell(CellRef::new(0, 1)).expect("B1");
    assert_eq!(cell_b1.formula.as_deref(), Some("SUM(A1)"));
}
