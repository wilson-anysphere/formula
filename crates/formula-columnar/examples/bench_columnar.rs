use formula_columnar::{
    ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions, Value,
};
use std::sync::Arc;
use std::time::Instant;

fn main() {
    let rows: usize = std::env::var("ROWS")
        .ok()
        .and_then(|v| v.parse().ok())
        .unwrap_or(1_000_000);

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

    let start_build = Instant::now();
    let mut builder = ColumnarTableBuilder::new(schema, options);
    for i in 0..rows {
        builder.append_row(&[
            Value::DateTime(i as i64),
            Value::String(categories[i % categories.len()].clone()),
            Value::Boolean(i % 2 == 0),
        ]);
    }
    let table = builder.finalize();
    let build_time = start_build.elapsed();

    println!("rows: {}", table.row_count());
    println!("compressed_bytes: {}", table.compressed_size_bytes());
    println!("build_time: {:?}", build_time);

    // Scan benchmark.
    let start_scan = Instant::now();
    let sum = table.scan().sum_f64(0).unwrap_or(0.0);
    let scan_time = start_scan.elapsed();
    println!("scan_sum(id): {} in {:?}", sum, scan_time);

    // Viewport benchmark.
    let start_view = Instant::now();
    let viewport = table.get_range(5_000.min(rows), (5_000 + 100).min(rows), 0, 3);
    let view_time = start_view.elapsed();
    println!(
        "viewport: {}x{} (col-major), fetch_time: {:?}",
        viewport.rows(),
        viewport.cols(),
        view_time
    );
}
