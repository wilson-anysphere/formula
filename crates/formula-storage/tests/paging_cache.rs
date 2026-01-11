use formula_storage::{
    CellChange, CellData, CellRange, CellValue, MemoryManager, MemoryManagerConfig, Storage,
};
use rusqlite::{Connection, OpenFlags};
use serde_json::json;

fn seed_dense_pages(
    storage: &Storage,
    sheet_id: uuid::Uuid,
    pages: usize,
    rows_per_page: usize,
    cols: i64,
) {
    let mut changes = Vec::new();
    for page in 0..pages {
        let row_base = (page * rows_per_page) as i64;
        for r in row_base..row_base + 10 {
            for c in 0..cols {
                changes.push(CellChange {
                    sheet_id,
                    row: r,
                    col: c,
                    data: CellData {
                        value: CellValue::Number(page as f64),
                        formula: None,
                        style: None,
                    },
                    user_id: None,
                });
            }
        }
    }
    storage
        .apply_cell_changes(&changes)
        .expect("seed workbook with cells");
}

#[test]
fn viewport_scrolling_evicts_pages_and_bounds_memory() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet", 0, None)
        .expect("create sheet");

    let rows_per_page = 50;
    let cols_per_page = 50;
    let page_count = 30;
    seed_dense_pages(&storage, sheet.id, page_count, rows_per_page, 10);

    let config = MemoryManagerConfig {
        max_memory_bytes: 20_000,
        max_pages: 128,
        eviction_watermark: 0.5,
        rows_per_page,
        cols_per_page,
    };
    let eviction_limit = (config.max_memory_bytes as f64 * config.eviction_watermark) as usize;
    let memory = MemoryManager::new(storage.clone(), config);

    for page in 0..page_count {
        let row_start = (page * rows_per_page) as i64;
        let viewport = CellRange::new(row_start, row_start + 20, 0, 20);
        let data = memory.load_viewport(sheet.id, viewport).expect("load viewport");
        assert!(
            data.get(row_start, 0).is_some(),
            "expected seeded cell in viewport"
        );

        let usage = memory.estimated_usage_bytes();
        assert!(
            usage <= eviction_limit,
            "cache usage {usage} exceeded eviction limit {eviction_limit}"
        );
    }

    let stats = memory.stats_snapshot();
    assert!(stats.pages_loaded > 0, "should load at least one page");
    assert!(stats.pages_evicted > 0, "expected pages to be evicted");
}

#[test]
fn dirty_edits_survive_eviction_and_are_persisted() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet", 0, None)
        .expect("create sheet");

    let rows_per_page = 50;
    let cols_per_page = 50;
    let page_count = 12;
    seed_dense_pages(&storage, sheet.id, page_count, rows_per_page, 10);

    let memory = MemoryManager::new(
        storage.clone(),
        MemoryManagerConfig {
            max_memory_bytes: 20_000,
            max_pages: 128,
            eviction_watermark: 0.5,
            rows_per_page,
            cols_per_page,
        },
    );

    // Load + edit a cell in the first page.
    memory
        .load_viewport(sheet.id, CellRange::new(0, 20, 0, 20))
        .expect("load first viewport");
    memory
        .record_change(CellChange {
            sheet_id: sheet.id,
            row: 0,
            col: 0,
            data: CellData {
                value: CellValue::Number(999.0),
                formula: None,
                style: None,
            },
            user_id: None,
        })
        .expect("record change");

    // Scroll far enough that the edited page becomes LRU and is evicted.
    for page in 1..page_count {
        let row_start = (page * rows_per_page) as i64;
        memory
            .load_viewport(sheet.id, CellRange::new(row_start, row_start + 20, 0, 20))
            .expect("scroll viewport");
    }

    let stats = memory.stats_snapshot();
    assert!(stats.pages_evicted > 0, "expected some evictions");
    assert!(
        stats.flush_transactions > 0,
        "dirty eviction should trigger a flush"
    );

    let persisted = storage
        .load_cells_in_range(sheet.id, CellRange::new(0, 0, 0, 0))
        .expect("load persisted cell");
    assert_eq!(persisted.len(), 1);
    assert_eq!(persisted[0].0, (0, 0));
    assert_eq!(persisted[0].1.value, CellValue::Number(999.0));
}

#[test]
fn load_viewport_returns_data_even_when_viewport_exceeds_page_budget() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet", 0, None)
        .expect("create sheet");

    let rows_per_page = 10;
    let cols_per_page = 10;

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
                row: rows_per_page as i64,
                col: 0,
                data: CellData {
                    value: CellValue::Number(2.0),
                    formula: None,
                    style: None,
                },
                user_id: None,
            },
            CellChange {
                sheet_id: sheet.id,
                row: (2 * rows_per_page) as i64,
                col: 0,
                data: CellData {
                    value: CellValue::Number(3.0),
                    formula: None,
                    style: None,
                },
                user_id: None,
            },
        ])
        .expect("seed cells");

    // Only allow 2 pages in cache, but request a viewport that spans 3 pages.
    let memory = MemoryManager::new(
        storage.clone(),
        MemoryManagerConfig {
            max_memory_bytes: 1024 * 1024,
            max_pages: 2,
            eviction_watermark: 1.0,
            rows_per_page,
            cols_per_page,
        },
    );

    let viewport = CellRange::new(0, (2 * rows_per_page) as i64, 0, 0);
    let data = memory.load_viewport(sheet.id, viewport).expect("load viewport");
    assert_eq!(
        data.get(0, 0).expect("cell 0,0").value,
        CellValue::Number(1.0)
    );
    assert_eq!(
        data.get(rows_per_page as i64, 0)
            .expect("cell 10,0")
            .value,
        CellValue::Number(2.0)
    );
    assert_eq!(
        data.get((2 * rows_per_page) as i64, 0)
            .expect("cell 20,0")
            .value,
        CellValue::Number(3.0)
    );

    assert!(
        memory.cached_page_count() <= 2,
        "cache should respect max_pages"
    );
}

#[test]
fn eviction_flush_preserves_global_change_order() {
    let uri = "file:paging_order?mode=memory&cache=shared";
    let storage = Storage::open_uri(uri).expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet", 0, None)
        .expect("create sheet");

    let rows_per_page = 10usize;
    let cols_per_page = 10usize;
    let memory = MemoryManager::new(
        storage.clone(),
        MemoryManagerConfig {
            max_memory_bytes: 1024 * 1024,
            max_pages: 2,
            eviction_watermark: 1.0,
            rows_per_page,
            cols_per_page,
        },
    );

    memory
        .record_change(CellChange {
            sheet_id: sheet.id,
            row: 0,
            col: 0,
            data: CellData {
                value: CellValue::Number(1.0),
                formula: None,
                style: None,
            },
            user_id: None,
        })
        .expect("change A");

    let row_b = rows_per_page as i64;
    memory
        .record_change(CellChange {
            sheet_id: sheet.id,
            row: row_b,
            col: 0,
            data: CellData {
                value: CellValue::Number(2.0),
                formula: None,
                style: None,
            },
            user_id: None,
        })
        .expect("change B");

    memory
        .record_change(CellChange {
            sheet_id: sheet.id,
            row: 0,
            col: 1,
            data: CellData {
                value: CellValue::Number(3.0),
                formula: None,
                style: None,
            },
            user_id: None,
        })
        .expect("change C");

    // Touch page 1 so page 0 becomes LRU, then load a third page to force an eviction.
    memory
        .load_viewport(sheet.id, CellRange::new(row_b, row_b, 0, 0))
        .expect("touch page 1");
    let row_c = (2 * rows_per_page) as i64;
    memory
        .load_viewport(sheet.id, CellRange::new(row_c, row_c, 0, 0))
        .expect("load page 2");

    let stats = memory.stats_snapshot();
    assert!(stats.flush_transactions > 0, "expected eviction flush");
    assert_eq!(stats.cell_changes_flushed, 3, "expected three flushed edits");

    let flags = OpenFlags::SQLITE_OPEN_READ_WRITE
        | OpenFlags::SQLITE_OPEN_CREATE
        | OpenFlags::SQLITE_OPEN_URI;
    let conn = Connection::open_with_flags(uri, flags).expect("open raw sqlite connection");
    let mut stmt = conn
        .prepare("SELECT target FROM change_log WHERE sheet_id = ?1 ORDER BY id")
        .expect("prepare change log query");
    let mut rows = stmt
        .query(rusqlite::params![sheet.id.to_string()])
        .expect("query change log");

    let mut targets = Vec::new();
    while let Some(row) = rows.next().expect("change log row") {
        let target: serde_json::Value = row.get(0).expect("target json");
        targets.push(target);
    }

    assert_eq!(
        targets,
        vec![
            json!({"row": 0, "col": 0}),
            json!({"row": row_b, "col": 0}),
            json!({"row": 0, "col": 1}),
        ]
    );
}

#[test]
fn load_viewport_with_margin_prefetches_neighbor_pages() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet", 0, None)
        .expect("create sheet");

    let rows_per_page = 10usize;
    let cols_per_page = 10usize;

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
                row: rows_per_page as i64,
                col: 0,
                data: CellData {
                    value: CellValue::Number(2.0),
                    formula: None,
                    style: None,
                },
                user_id: None,
            },
        ])
        .expect("seed cells");

    let memory = MemoryManager::new(
        storage,
        MemoryManagerConfig {
            max_memory_bytes: 1024 * 1024,
            max_pages: 64,
            eviction_watermark: 1.0,
            rows_per_page,
            cols_per_page,
        },
    );

    // Load row 0 but prefetch an extra page of rows so row 10 is loaded too.
    memory
        .load_viewport_with_margin(
            sheet.id,
            CellRange::new(0, 0, 0, 0),
            rows_per_page as i64,
            0,
        )
        .expect("load viewport with margin");

    assert!(
        memory.get_cached_cell(sheet.id, rows_per_page as i64, 0).is_some(),
        "expected neighboring page to be prefetched into cache"
    );

    let before = memory.stats_snapshot();
    memory
        .load_viewport(sheet.id, CellRange::new(rows_per_page as i64, rows_per_page as i64, 0, 0))
        .expect("load second viewport");
    let after = memory.stats_snapshot();

    assert!(
        after.page_hits > before.page_hits,
        "expected adjacent viewport to be a cache hit after prefetch"
    );
    assert_eq!(
        after.page_misses, before.page_misses,
        "expected no cache miss after prefetch"
    );
}
