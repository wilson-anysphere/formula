//! A small, dependency-free micro-benchmark comparing immutable build vs incremental refresh.
//!
//! Usage:
//! ```bash
//! # Defaults to 1_000_000 base rows + 100_000 appended rows.
//! cargo run -p formula-columnar --example bench_refresh_compare --release
//!
//! # Customize sizes.
//! BASE_ROWS=200000 APPEND_ROWS=20000 cargo run -p formula-columnar --example bench_refresh_compare --release
//! ```

use formula_columnar::{
    ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions, Value,
};
use std::sync::Arc;
use std::time::Instant;

fn main() {
    let base_rows: usize = std::env::var("BASE_ROWS")
        .ok()
        .and_then(|v| v.parse().ok())
        .unwrap_or(1_000_000);
    let append_rows: usize = std::env::var("APPEND_ROWS")
        .ok()
        .and_then(|v| v.parse().ok())
        .unwrap_or(100_000);

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
        page_size_rows: 65_536,
        cache: PageCacheConfig { max_entries: 32 },
    };

    let categories: Vec<Arc<str>> = ["AA", "BB", "CC", "DD", "EE", "FF", "GG", "HH", "II", "JJ"]
        .into_iter()
        .map(Arc::<str>::from)
        .collect();

    println!(
        "base_rows={} append_rows={} page_size_rows={}",
        base_rows, append_rows, options.page_size_rows
    );

    // -------------------------------------------------------------------------
    // Baseline: build immutable table from scratch (base + appended).
    // -------------------------------------------------------------------------
    let start_baseline = Instant::now();
    let mut builder = ColumnarTableBuilder::new(schema.clone(), options);
    for i in 0..(base_rows + append_rows) {
        builder.append_row(&[
            Value::DateTime(i as i64),
            Value::String(categories[i % categories.len()].clone()),
            Value::Boolean(i % 2 == 0),
        ]);
    }
    let baseline_table = builder.finalize();
    let baseline_build = start_baseline.elapsed();

    let start_baseline_range = Instant::now();
    let _ = baseline_table.get_range(5_000.min(baseline_table.row_count()), 5_100.min(baseline_table.row_count()), 0, 3);
    let baseline_range = start_baseline_range.elapsed();

    println!(
        "[baseline] build={:?} range(100x3)={:?} compressed_bytes={}",
        baseline_build,
        baseline_range,
        baseline_table.compressed_size_bytes()
    );

    // -------------------------------------------------------------------------
    // Incremental refresh: build base once, then append + freeze.
    // -------------------------------------------------------------------------
    let start_base = Instant::now();
    let mut base_builder = ColumnarTableBuilder::new(schema.clone(), options);
    for i in 0..base_rows {
        base_builder.append_row(&[
            Value::DateTime(i as i64),
            Value::String(categories[i % categories.len()].clone()),
            Value::Boolean(i % 2 == 0),
        ]);
    }
    let base_table = base_builder.finalize();
    let base_build = start_base.elapsed();

    let start_append = Instant::now();
    let mut mutable = base_table.into_mutable();
    for i in 0..append_rows {
        let id = base_rows + i;
        mutable.append_row(&[
            Value::DateTime(id as i64),
            Value::String(categories[id % categories.len()].clone()),
            Value::Boolean(id % 2 == 0),
        ]);
    }
    let append_time = start_append.elapsed();

    let start_freeze = Instant::now();
    let refreshed = mutable.freeze();
    let freeze_time = start_freeze.elapsed();

    let start_refreshed_range = Instant::now();
    let _ = refreshed.get_range(5_000.min(refreshed.row_count()), 5_100.min(refreshed.row_count()), 0, 3);
    let refreshed_range = start_refreshed_range.elapsed();

    println!(
        "[incremental] base_build={:?} append={:?} freeze={:?} range(100x3)={:?} compressed_bytes={}",
        base_build,
        append_time,
        freeze_time,
        refreshed_range,
        refreshed.compressed_size_bytes()
    );

    // A loose sanity ratio printout for manual inspection.
    if baseline_build.as_nanos() > 0 {
        let ratio = (append_time.as_secs_f64() + freeze_time.as_secs_f64()) / baseline_build.as_secs_f64();
        println!("[ratio] (append+freeze)/baseline_build = {:.3}", ratio);
    }
}
