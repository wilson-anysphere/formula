use formula_storage::{CellChange, CellData, CellRange, CellValue, MemoryManager, MemoryManagerConfig, Storage};

#[test]
fn bench_load_viewport_hot_cache() {
    if std::env::var("FORMULA_STORAGE_BENCH").is_err() {
        return;
    }

    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet", 0, None)
        .expect("create sheet");

    // Seed a modest block of cells so each viewport has work to do.
    let mut changes = Vec::new();
    for row in 0..200i64 {
        for col in 0..50i64 {
            changes.push(CellChange {
                sheet_id: sheet.id,
                row,
                col,
                data: CellData {
                    value: CellValue::Number((row * 100 + col) as f64),
                    formula: None,
                    style: None,
                },
                user_id: None,
            });
        }
    }
    storage.apply_cell_changes(&changes).expect("seed");

    let memory = MemoryManager::new(
        storage,
        MemoryManagerConfig {
            max_memory_bytes: 32 * 1024 * 1024,
            max_pages: 4096,
            eviction_watermark: 0.9,
            rows_per_page: 64,
            cols_per_page: 64,
        },
    );

    let viewport = CellRange::new(0, 120, 0, 40);
    memory
        .load_viewport(sheet.id, viewport)
        .expect("warm viewport");

    let iters = 5_000usize;
    let start = std::time::Instant::now();
    for _ in 0..iters {
        let _ = memory
            .load_viewport(sheet.id, viewport)
            .expect("hot cache viewport");
    }
    let elapsed = start.elapsed();
    eprintln!(
        "bench_load_viewport_hot_cache: {iters} iters in {:?} ({:?}/iter)",
        elapsed,
        elapsed / iters as u32
    );
}

