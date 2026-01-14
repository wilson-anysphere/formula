//! Large dataset integration tests.
//!
//! This file intentionally contains two tiers:
//! - A **smoke test** (runs by default) that builds a moderately large columnar table and
//!   fetches a small viewport via [`Worksheet::get_range_batch`]. This ensures the
//!   columnar-backed worksheet viewport path is exercised in CI / normal `cargo test` runs.
//! - A **stress test** (ignored by default) that scales up to ~10M rows to catch OOM / extreme
//!   performance regressions.
//!
//! To run the ignored stress test via the standard agent wrapper:
//! `bash -lc '. scripts/agent-init.sh && bash scripts/cargo_agent.sh test -p formula-model --test large_dataset -- --ignored'`
use formula_columnar::{
    ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions, Value,
};
use formula_model::{CellRef, CellValue, Range, Worksheet};
use std::sync::Arc;

fn build_table(rows: usize) -> formula_columnar::ColumnarTable {
    let schema = vec![
        ColumnSchema {
            name: "id".to_owned(),
            column_type: ColumnType::DateTime,
        },
        ColumnSchema {
            name: "category".to_owned(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "flag".to_owned(),
            column_type: ColumnType::Boolean,
        },
    ];

    let options = TableOptions {
        // Match the import defaults; also useful to exercise crossing a page boundary.
        page_size_rows: 65_536,
        cache: PageCacheConfig { max_entries: 32 },
    };

    let mut builder = ColumnarTableBuilder::new(schema, options);
    let cats: Vec<Arc<str>> = ["AA", "BB", "CC", "DD", "EE", "FF", "GG", "HH", "II", "JJ"]
        .into_iter()
        .map(Arc::<str>::from)
        .collect();
    for i in 0..rows {
        builder.append_row(&[
            Value::DateTime(i as i64),
            Value::String(cats[i % cats.len()].clone()),
            Value::Boolean(i % 2 == 0),
        ]);
    }

    builder.finalize()
}

#[test]
fn stream_100k_rows_and_render_viewport_smoke() {
    // This is a CI-friendly default that still ensures `get_range_batch` doesn't scale with
    // the total table size when reading a small viewport.
    let rows: usize = std::env::var("FORMULA_LARGE_TEST_ROWS_SMOKE")
        .ok()
        .and_then(|v| v.parse().ok())
        .unwrap_or(100_000);

    let table = build_table(rows);
    let mut sheet = Worksheet::new(1, "Data");
    sheet.set_columnar_table(CellRef::new(0, 0), Arc::new(table));

    // Simulate UI requesting a typical visible viewport (100x3), chosen to straddle a page
    // boundary when using a 65_536-row page size.
    const VIEWPORT_ROWS: u32 = 100;
    assert!(
        rows >= VIEWPORT_ROWS as usize,
        "smoke dataset too small for 100-row viewport: rows={rows}"
    );
    let start_row = if rows >= 65_500 + VIEWPORT_ROWS as usize {
        65_500u32
    } else {
        ((rows - VIEWPORT_ROWS as usize) / 2) as u32
    };
    let range = Range::new(
        CellRef::new(start_row, 0),
        CellRef::new(start_row + VIEWPORT_ROWS - 1, 2),
    );
    let viewport = sheet.get_range_batch(range);

    // Guardrails: dimensions must match the requested viewport (not the backing table size).
    assert_eq!(viewport.start, range.start);
    assert_eq!(viewport.columns.len(), 3);
    assert!(viewport
        .columns
        .iter()
        .all(|c| c.len() == VIEWPORT_ROWS as usize));
    assert_eq!(
        sheet.cell_count(),
        0,
        "reading a viewport must not populate sparse cells"
    );

    // Spot-check a few representative values.
    const CATS: [&str; 10] = ["AA", "BB", "CC", "DD", "EE", "FF", "GG", "HH", "II", "JJ"];
    let cat_at = |row: u32| CATS[row as usize % CATS.len()];
    let flag_at = |row: u32| row % 2 == 0;

    assert_eq!(viewport.columns[0][0], CellValue::Number(start_row as f64));
    assert_eq!(
        viewport.columns[1][0],
        CellValue::String(cat_at(start_row).to_string())
    );
    assert_eq!(
        viewport.columns[2][0],
        CellValue::Boolean(flag_at(start_row))
    );

    // Spot-check a row well inside the viewport. When `start_row` is 65_500 (the default),
    // this crosses the 65_536 row page boundary.
    assert_eq!(
        viewport.columns[0][37],
        CellValue::Number((start_row + 37) as f64)
    );
    assert_eq!(
        viewport.columns[1][37],
        CellValue::String(cat_at(start_row + 37).to_string())
    );
    assert_eq!(
        viewport.columns[2][37],
        CellValue::Boolean(flag_at(start_row + 37))
    );

    assert_eq!(
        viewport.columns[0][99],
        CellValue::Number((start_row + VIEWPORT_ROWS - 1) as f64)
    );
    assert_eq!(
        viewport.columns[1][99],
        CellValue::String(cat_at(start_row + VIEWPORT_ROWS - 1).to_string())
    );
    assert_eq!(
        viewport.columns[2][99],
        CellValue::Boolean(flag_at(start_row + VIEWPORT_ROWS - 1))
    );
}

#[test]
#[ignore]
/// Stress test: stream a very large dataset and render a small viewport.
///
/// This is ignored by default because it can take a while in debug mode.
/// Run with:
/// `bash -lc '. scripts/agent-init.sh && bash scripts/cargo_agent.sh test -p formula-model --test large_dataset -- --ignored'`
fn stream_10m_rows_and_render_first_viewport_without_oom() {
    let rows: usize = std::env::var("FORMULA_LARGE_TEST_ROWS_STRESS")
        .ok()
        .or_else(|| std::env::var("FORMULA_LARGE_TEST_ROWS").ok())
        .and_then(|v| v.parse().ok())
        .unwrap_or(10_000_000);

    let table = build_table(rows);
    let mut sheet = Worksheet::new(1, "Data");
    sheet.set_columnar_table(CellRef::new(0, 0), Arc::new(table));

    // Simulate UI requesting the first visible viewport.
    let viewport = sheet.get_range_batch(Range::new(CellRef::new(0, 0), CellRef::new(99, 2)));
    assert_eq!(viewport.columns.len(), 3);
    assert_eq!(viewport.columns[0][0], CellValue::Number(0.0));
    assert_eq!(viewport.columns[1][0], CellValue::String("AA".to_string()));
    assert_eq!(viewport.columns[2][0], CellValue::Boolean(true));
}
