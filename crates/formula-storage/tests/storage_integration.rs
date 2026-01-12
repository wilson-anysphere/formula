use formula_model::{
    ArrayValue, CellRef, ErrorValue, RichText, SheetNameError, SpillValue, TabColor,
    EXCEL_MAX_SHEET_NAME_LEN,
};
use formula_storage::storage::StorageError;
use formula_storage::{
    AutoSaveConfig, AutoSaveManager, CellChange, CellData, CellRange, CellValue, SheetVisibility,
    MemoryManager, MemoryManagerConfig, NamedRange, Storage, Style,
};
use rusqlite::{Connection, OpenFlags};
use serde_json::json;
use std::time::Duration;
use tempfile::NamedTempFile;
use uuid::Uuid;

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
                    value: CellValue::String("hello".to_string()),
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
    assert_eq!(cells[1].1.value, CellValue::String("hello".to_string()));

    // Keep storage1 alive for the lifetime of the test to ensure the shared
    // in-memory DB isn't dropped.
    std::mem::drop(storage1);
}

#[test]
fn sparse_storage_only_persists_non_empty_cells() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
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
        alignment: Some(json!({"outer": {"b": 2, "a": 1}, "a": 0})),
        protection: None,
    };
    let style_b = Style {
        font_id: None,
        fill_id: None,
        border_id: None,
        number_format: Some("0.0".to_string()),
        alignment: Some(json!({"a": 0, "outer": {"a": 1, "b": 2}})),
        protection: None,
    };

    let id_a = storage.get_or_insert_style(&style_a).expect("insert style");
    let id_b = storage.get_or_insert_style(&style_b).expect("dedup style");
    assert_eq!(id_a, id_b);

    let style_c = Style {
        number_format: Some("0.00".to_string()),
        ..style_a
    };
    let id_c = storage
        .get_or_insert_style(&style_c)
        .expect("insert other style");
    assert_ne!(id_a, id_c);
}

#[test]
fn apply_cell_changes_preserves_style_when_style_is_omitted() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet", 0, None)
        .expect("create sheet");

    let style = Style {
        font_id: None,
        fill_id: None,
        border_id: None,
        number_format: Some("0.00".to_string()),
        alignment: None,
        protection: None,
    };

    storage
        .apply_cell_changes(&[CellChange {
            sheet_id: sheet.id,
            row: 0,
            col: 0,
            data: CellData {
                value: CellValue::Number(1.0),
                formula: None,
                style: Some(style),
            },
            user_id: None,
        }])
        .expect("set styled cell");

    let initial = storage
        .load_cells_in_range(sheet.id, CellRange::new(0, 0, 0, 0))
        .expect("load initial cell");
    let initial_style_id = initial
        .first()
        .map(|(_, snap)| snap.style_id)
        .flatten()
        .expect("expected initial style id");

    storage
        .apply_cell_changes(&[CellChange {
            sheet_id: sheet.id,
            row: 0,
            col: 0,
            data: CellData {
                value: CellValue::Number(2.0),
                formula: None,
                style: None,
            },
            user_id: None,
        }])
        .expect("update cell without style payload");

    let updated = storage
        .load_cells_in_range(sheet.id, CellRange::new(0, 0, 0, 0))
        .expect("load updated cell");
    let updated_style_id = updated
        .first()
        .map(|(_, snap)| snap.style_id)
        .flatten()
        .expect("expected updated style id");
    assert_eq!(updated_style_id, initial_style_id);

    // Clearing the cell contents should keep the style row so the cell remains formatted.
    storage
        .apply_cell_changes(&[CellChange {
            sheet_id: sheet.id,
            row: 0,
            col: 0,
            data: CellData::empty(),
            user_id: None,
        }])
        .expect("clear cell");

    let cleared = storage
        .load_cells_in_range(sheet.id, CellRange::new(0, 0, 0, 0))
        .expect("load cleared cell");
    assert_eq!(cleared.len(), 1, "expected style-only cell to remain");
    let cleared_snap = &cleared[0].1;
    assert!(cleared_snap.value.is_empty());
    assert!(cleared_snap.formula.is_none());
    assert_eq!(cleared_snap.style_id, Some(initial_style_id));
}

#[test]
fn change_log_records_cell_operations() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet", 0, None)
        .expect("create sheet");

    storage
        .apply_cell_changes(&[CellChange {
            sheet_id: sheet.id,
            row: 0,
            col: 0,
            data: CellData {
                value: CellValue::String("hello".to_string()),
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
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet", 0, None)
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
                value: CellValue::Number(10.0),
                formula: None,
                style: None,
            },
            user_id: None,
        })
        .expect("record change");

    autosave
        .record_change(CellChange {
            sheet_id: sheet.id,
            row: 0,
            col: 1,
            data: CellData {
                value: CellValue::Number(20.0),
                formula: None,
                style: None,
            },
            user_id: None,
        })
        .expect("record change");

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

#[test]
fn rich_cell_values_round_trip() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet", 0, None)
        .expect("create sheet");

    let rich = RichText::from_segments(vec![
        ("Hello ".to_string(), Default::default()),
        (
            "World".to_string(),
            formula_model::rich_text::RichTextRunStyle {
                bold: Some(true),
                ..Default::default()
            },
        ),
    ]);

    let array = ArrayValue {
        data: vec![
            vec![CellValue::Number(1.0), CellValue::String("x".to_string())],
            vec![CellValue::Boolean(true), CellValue::Error(ErrorValue::NA)],
        ],
    };

    storage
        .apply_cell_changes(&[
            CellChange {
                sheet_id: sheet.id,
                row: 0,
                col: 0,
                data: CellData {
                    value: CellValue::RichText(rich.clone()),
                    formula: None,
                    style: None,
                },
                user_id: None,
            },
            CellChange {
                sheet_id: sheet.id,
                row: 1,
                col: 0,
                data: CellData {
                    value: CellValue::Array(array.clone()),
                    formula: None,
                    style: None,
                },
                user_id: None,
            },
            CellChange {
                sheet_id: sheet.id,
                row: 2,
                col: 0,
                data: CellData {
                    value: CellValue::Spill(SpillValue {
                        origin: CellRef::new(1, 0),
                    }),
                    formula: None,
                    style: None,
                },
                user_id: None,
            },
            CellChange {
                sheet_id: sheet.id,
                row: 3,
                col: 0,
                data: CellData {
                    value: CellValue::Error(ErrorValue::Div0),
                    formula: None,
                    style: None,
                },
                user_id: None,
            },
        ])
        .expect("persist cells");

    let cells = storage
        .load_cells_in_range(sheet.id, CellRange::new(0, 10, 0, 10))
        .expect("load cells");

    let mut by_coord = std::collections::HashMap::new();
    for (coord, snap) in cells {
        by_coord.insert(coord, snap.value);
    }

    assert_eq!(by_coord.get(&(0, 0)), Some(&CellValue::RichText(rich)));
    assert_eq!(by_coord.get(&(1, 0)), Some(&CellValue::Array(array)));
    assert_eq!(
        by_coord.get(&(2, 0)),
        Some(&CellValue::Spill(SpillValue {
            origin: CellRef::new(1, 0)
        }))
    );
    assert_eq!(
        by_coord.get(&(3, 0)),
        Some(&CellValue::Error(ErrorValue::Div0))
    );
}

#[test]
fn opens_and_migrates_legacy_schema() {
    let tmp = NamedTempFile::new().expect("tmpfile");
    let path = tmp.path();

    // Simulate a database created by the pre-versioned schema (no schema_version, no value_json).
    let conn = Connection::open(path).expect("open legacy db");
    conn.execute_batch(
        r#"
        CREATE TABLE IF NOT EXISTS workbooks (
          id TEXT PRIMARY KEY,
          name TEXT NOT NULL,
          created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
          modified_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
          metadata JSON
        );

        CREATE TABLE IF NOT EXISTS sheets (
          id TEXT PRIMARY KEY,
          workbook_id TEXT REFERENCES workbooks(id),
          name TEXT NOT NULL,
          position INTEGER,
          visibility TEXT NOT NULL DEFAULT 'visible' CHECK (visibility IN ('visible','hidden','veryHidden')),
          tab_color TEXT,
          xlsx_sheet_id INTEGER,
          xlsx_rel_id TEXT,
          frozen_rows INTEGER DEFAULT 0,
          frozen_cols INTEGER DEFAULT 0,
          zoom REAL DEFAULT 1.0,
          metadata JSON
        );

        CREATE TABLE IF NOT EXISTS cells (
          sheet_id TEXT REFERENCES sheets(id),
          row INTEGER,
          col INTEGER,
          value_type TEXT,
          value_number REAL,
          value_string TEXT,
          formula TEXT,
          style_id INTEGER,
          PRIMARY KEY (sheet_id, row, col)
        );
        "#,
    )
    .expect("create legacy schema");

    let workbook_id = Uuid::parse_str("00000000-0000-0000-0000-000000000001").unwrap();
    let sheet_id = Uuid::parse_str("00000000-0000-0000-0000-000000000002").unwrap();

    conn.execute(
        "INSERT INTO workbooks (id, name) VALUES (?1, ?2)",
        rusqlite::params![workbook_id.to_string(), "Book"],
    )
    .expect("insert workbook");
    conn.execute(
        "INSERT INTO sheets (id, workbook_id, name, position) VALUES (?1, ?2, ?3, ?4)",
        rusqlite::params![
            sheet_id.to_string(),
            workbook_id.to_string(),
            "Sheet1",
            0i64
        ],
    )
    .expect("insert sheet");
    conn.execute(
        "INSERT INTO cells (sheet_id, row, col, value_type, value_string) VALUES (?1, ?2, ?3, ?4, ?5)",
        rusqlite::params![sheet_id.to_string(), 0i64, 0i64, "error", "#DIV/0!"],
    )
    .expect("insert cell");
    drop(conn);

    let storage = Storage::open_path(path).expect("open with migration");
    let cells = storage
        .load_cells_in_range(sheet_id, CellRange::new(0, 0, 0, 0))
        .expect("load cells");
    assert_eq!(cells.len(), 1);
    assert_eq!(cells[0].1.value, CellValue::Error(ErrorValue::Div0));
    drop(storage);

    // Confirm the migration added the new column.
    let conn = Connection::open(path).expect("reopen");
    let mut stmt = conn.prepare("PRAGMA table_info(cells)").expect("pragma");
    let cols = stmt
        .query_map([], |row| row.get::<_, String>(1))
        .expect("query pragma")
        .collect::<std::result::Result<Vec<_>, _>>()
        .expect("collect pragma");
    assert!(
        cols.iter().any(|c| c == "value_json"),
        "value_json column missing after migration"
    );

    let mut stmt = conn.prepare("PRAGMA table_info(sheets)").expect("pragma sheets");
    let sheet_cols = stmt
        .query_map([], |row| row.get::<_, String>(1))
        .expect("query pragma sheets")
        .collect::<std::result::Result<Vec<_>, _>>()
        .expect("collect pragma sheets");
    for required in ["model_sheet_id", "tab_color_json", "model_sheet_json"] {
        assert!(
            sheet_cols.iter().any(|c| c == required),
            "{required} column missing after migration"
        );
    }
}

#[test]
fn sheet_metadata_persists_visibility_tab_color_and_xlsx_ids() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet");

    storage
        .set_sheet_visibility(sheet.id, SheetVisibility::Hidden)
        .expect("set visibility");
    let tab_color = TabColor::rgb("FFFF0000");
    storage
        .set_sheet_tab_color(sheet.id, Some(&tab_color))
        .expect("set tab color");
    storage
        .set_sheet_xlsx_metadata(sheet.id, Some(42), Some("rId7"))
        .expect("set xlsx metadata");
    storage.rename_sheet(sheet.id, "Renamed").expect("rename");

    let loaded = storage.get_sheet_meta(sheet.id).expect("get sheet");
    assert_eq!(loaded.name, "Renamed");
    assert_eq!(loaded.visibility, SheetVisibility::Hidden);
    assert_eq!(loaded.tab_color.as_deref(), Some("FFFF0000"));
    assert_eq!(loaded.xlsx_sheet_id, Some(42));
    assert_eq!(loaded.xlsx_rel_id.as_deref(), Some("rId7"));
}

#[test]
fn sheet_metadata_persists_non_rgb_tab_color_via_model_export() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet");

    let tab_color = TabColor {
        theme: Some(3),
        tint: Some(0.25),
        ..Default::default()
    };
    storage
        .set_sheet_tab_color(sheet.id, Some(&tab_color))
        .expect("set tab color");

    let exported = storage
        .export_model_workbook(workbook.id)
        .expect("export workbook");
    let sheet = exported
        .sheets
        .iter()
        .find(|s| s.name == "Sheet1")
        .expect("Sheet1 exists");
    assert_eq!(sheet.tab_color, Some(tab_color));
}

#[test]
fn sheet_reorder_and_delete_renormalize_positions() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet_a = storage
        .create_sheet(workbook.id, "SheetA", 0, None)
        .expect("create sheet A");
    let _sheet_b = storage
        .create_sheet(workbook.id, "SheetB", 1, None)
        .expect("create sheet B");
    let sheet_c = storage
        .create_sheet(workbook.id, "SheetC", 2, None)
        .expect("create sheet C");

    storage.reorder_sheet(sheet_c.id, 0).expect("reorder");
    let sheets = storage.list_sheets(workbook.id).expect("list sheets");
    assert_eq!(
        sheets.iter().map(|s| s.name.as_str()).collect::<Vec<_>>(),
        vec!["SheetC", "SheetA", "SheetB"]
    );
    assert_eq!(
        sheets.iter().map(|s| s.position).collect::<Vec<_>>(),
        vec![0, 1, 2]
    );

    storage.delete_sheet(sheet_a.id).expect("delete");
    let sheets = storage.list_sheets(workbook.id).expect("list after delete");
    assert_eq!(
        sheets.iter().map(|s| s.name.as_str()).collect::<Vec<_>>(),
        vec!["SheetC", "SheetB"]
    );
    assert_eq!(
        sheets.iter().map(|s| s.position).collect::<Vec<_>>(),
        vec![0, 1]
    );
}

#[test]
fn sheet_reorder_sheets_batch_updates_positions() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet_a = storage
        .create_sheet(workbook.id, "SheetA", 0, None)
        .expect("create sheet A");
    let sheet_b = storage
        .create_sheet(workbook.id, "SheetB", 1, None)
        .expect("create sheet B");
    let sheet_c = storage
        .create_sheet(workbook.id, "SheetC", 2, None)
        .expect("create sheet C");

    storage
        .reorder_sheets(workbook.id, &[sheet_c.id, sheet_a.id, sheet_b.id])
        .expect("reorder");
    let sheets = storage.list_sheets(workbook.id).expect("list sheets");
    assert_eq!(
        sheets.iter().map(|s| s.name.as_str()).collect::<Vec<_>>(),
        vec!["SheetC", "SheetA", "SheetB"]
    );
    assert_eq!(
        sheets.iter().map(|s| s.position).collect::<Vec<_>>(),
        vec![0, 1, 2]
    );
}

#[test]
fn create_sheet_inserts_at_position_and_renormalizes_positions() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    storage
        .create_sheet(workbook.id, "SheetA", 0, None)
        .expect("create sheet A");
    storage
        .create_sheet(workbook.id, "SheetB", 1, None)
        .expect("create sheet B");
    storage
        .create_sheet(workbook.id, "SheetC", 2, None)
        .expect("create sheet C");

    let inserted = storage
        .create_sheet(workbook.id, "Inserted", 1, None)
        .expect("create inserted sheet");
    assert_eq!(inserted.position, 1);

    let sheets = storage.list_sheets(workbook.id).expect("list sheets");
    assert_eq!(
        sheets.iter().map(|s| s.name.as_str()).collect::<Vec<_>>(),
        vec!["SheetA", "Inserted", "SheetB", "SheetC"]
    );
    assert_eq!(
        sheets.iter().map(|s| s.position).collect::<Vec<_>>(),
        vec![0, 1, 2, 3]
    );
}

#[test]
fn list_sheets_tolerates_null_positions() {
    use rusqlite::{Connection, OpenFlags};

    let uri = "file:sheet_null_positions?mode=memory&cache=shared";
    let storage = Storage::open_uri(uri).expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet_a = storage
        .create_sheet(workbook.id, "SheetA", 0, None)
        .expect("create sheet A");
    let sheet_b = storage
        .create_sheet(workbook.id, "SheetB", 1, None)
        .expect("create sheet B");

    let flags =
        OpenFlags::SQLITE_OPEN_READ_WRITE | OpenFlags::SQLITE_OPEN_CREATE | OpenFlags::SQLITE_OPEN_URI;
    let conn = Connection::open_with_flags(uri, flags).expect("open raw connection");
    conn.execute(
        "UPDATE sheets SET position = NULL WHERE workbook_id = ?1",
        rusqlite::params![workbook.id.to_string()],
    )
    .expect("null out sheet positions");

    // `list_sheets` should coalesce NULL positions so legacy/corrupt databases don't panic.
    let sheets = storage.list_sheets(workbook.id).expect("list sheets");
    assert_eq!(sheets.len(), 2);
    assert!(sheets.iter().all(|s| s.position == 0));

    let mut expected = vec![sheet_a.id.to_string(), sheet_b.id.to_string()];
    expected.sort();
    let actual = sheets.iter().map(|s| s.id.to_string()).collect::<Vec<_>>();
    assert_eq!(actual, expected);

    // Creating a sheet should renormalize positions back to a contiguous ordering.
    storage
        .create_sheet(workbook.id, "Inserted", 0, None)
        .expect("create sheet");
    let sheets = storage.list_sheets(workbook.id).expect("list sheets after insert");
    assert_eq!(
        sheets.iter().map(|s| s.position).collect::<Vec<_>>(),
        vec![0, 1, 2]
    );
}

#[test]
fn sheet_names_are_unique_case_insensitive() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    storage
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet");
    let err = storage
        .create_sheet(workbook.id, "sheet1", 1, None)
        .expect_err("duplicate");

    match err {
        StorageError::DuplicateSheetName(name) => assert_eq!(name, "sheet1"),
        other => panic!("unexpected error: {other:?}"),
    }
}

#[test]
fn sheet_names_match_excel_validation_rules() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");

    assert!(matches!(
        storage.create_sheet(workbook.id, "", 0, None),
        Err(StorageError::EmptySheetName)
    ));
    assert!(matches!(
        storage.create_sheet(workbook.id, "   ", 0, None),
        Err(StorageError::EmptySheetName)
    ));

    for ch in [':', '\\', '/', '?', '*', '[', ']'] {
        let storage = Storage::open_in_memory().expect("open storage");
        let workbook = storage
            .create_workbook("Book", None)
            .expect("create workbook");
        let name = format!("Bad{ch}Name");
        assert!(matches!(
            storage.create_sheet(workbook.id, &name, 0, None),
            Err(StorageError::InvalidSheetName(
                SheetNameError::InvalidCharacter(c)
            )) if c == ch
        ));
    }

    let max = "a".repeat(EXCEL_MAX_SHEET_NAME_LEN);
    storage
        .create_sheet(workbook.id, &max, 0, None)
        .expect("max length ok");

    let too_long = "a".repeat(EXCEL_MAX_SHEET_NAME_LEN + 1);
    assert!(matches!(
        storage.create_sheet(workbook.id, &too_long, 1, None),
        Err(StorageError::InvalidSheetName(SheetNameError::TooLong))
    ));

    assert!(matches!(
        storage.create_sheet(workbook.id, "'Leading", 2, None),
        Err(StorageError::InvalidSheetName(
            SheetNameError::LeadingOrTrailingApostrophe
        ))
    ));
    assert!(matches!(
        storage.create_sheet(workbook.id, "Trailing'", 3, None),
        Err(StorageError::InvalidSheetName(
            SheetNameError::LeadingOrTrailingApostrophe
        ))
    ));
}

#[test]
fn sheet_names_are_unique_unicode_case_insensitive() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    storage
        .create_sheet(workbook.id, "Äbc", 0, None)
        .expect("create sheet");

    let err = storage
        .create_sheet(workbook.id, "äbc", 1, None)
        .expect_err("duplicate");
    assert!(matches!(err, StorageError::DuplicateSheetName(name) if name == "äbc"));
}

#[test]
fn rename_sheet_validates_names_and_enforces_uniqueness() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let data = storage
        .create_sheet(workbook.id, "Data", 0, None)
        .expect("create sheet");
    let summary = storage
        .create_sheet(workbook.id, "Summary", 1, None)
        .expect("create sheet");

    let err = storage.rename_sheet(data.id, "Bad:Name").expect_err("invalid");
    assert!(matches!(
        err,
        StorageError::InvalidSheetName(SheetNameError::InvalidCharacter(':'))
    ));

    let err = storage.rename_sheet(summary.id, "DATA").expect_err("duplicate");
    assert!(matches!(err, StorageError::DuplicateSheetName(name) if name == "DATA"));

    // Renaming the same sheet to a different case is allowed.
    storage.rename_sheet(data.id, "data").expect("rename");
    let loaded = storage.get_sheet_meta(data.id).expect("get sheet");
    assert_eq!(loaded.name, "data");
}

#[test]
fn named_ranges_are_case_insensitive() {
    let uri = "file:named_ranges_ci?mode=memory&cache=shared";
    let storage = Storage::open_uri(uri).expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");

    storage
        .upsert_named_range(&NamedRange {
            workbook_id: workbook.id,
            name: "MyRange".to_string(),
            scope: "workbook".to_string(),
            reference: "Sheet1!$A$1".to_string(),
        })
        .expect("insert named range");

    let fetched = storage
        .get_named_range(workbook.id, "myrange", "WORKBOOK")
        .expect("get named range")
        .expect("named range exists");
    assert_eq!(fetched.reference, "Sheet1!$A$1");

    storage
        .upsert_named_range(&NamedRange {
            workbook_id: workbook.id,
            name: "MYRANGE".to_string(),
            scope: "workbook".to_string(),
            reference: "Sheet1!$B$2".to_string(),
        })
        .expect("update named range");

    let fetched = storage
        .get_named_range(workbook.id, "MyRange", "workbook")
        .expect("get named range")
        .expect("named range exists");
    assert_eq!(fetched.reference, "Sheet1!$B$2");

    let flags =
        OpenFlags::SQLITE_OPEN_READ_WRITE | OpenFlags::SQLITE_OPEN_CREATE | OpenFlags::SQLITE_OPEN_URI;
    let conn = Connection::open_with_flags(uri, flags).expect("open raw connection");
    let count: i64 = conn
        .query_row(
            "SELECT COUNT(*) FROM named_ranges WHERE workbook_id = ?1",
            rusqlite::params![workbook.id.to_string()],
            |r| r.get(0),
        )
        .expect("count named ranges");
    assert_eq!(count, 1);
}
