//! Large dataset integration test.
//!
//! This test is ignored by default because it can take a while in debug mode.
//! Run with:
//! `cargo test -p formula-model --test large_dataset -- --ignored`
use formula_columnar::{
    ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions, Value,
};
use formula_model::{CellRef, CellValue, Range, Worksheet};
use std::sync::Arc;

#[test]
#[ignore]
fn stream_10m_rows_and_render_first_viewport_without_oom() {
    let rows: usize = std::env::var("FORMULA_LARGE_TEST_ROWS")
        .ok()
        .and_then(|v| v.parse().ok())
        .unwrap_or(10_000_000);

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

    let table = builder.finalize();
    let mut sheet = Worksheet::new(1, "Data");
    sheet.set_columnar_table(CellRef::new(0, 0), Arc::new(table));

    // Simulate UI requesting the first visible viewport.
    let viewport = sheet.get_range_batch(Range::new(CellRef::new(0, 0), CellRef::new(99, 2)));
    assert_eq!(viewport.columns.len(), 3);
    assert_eq!(viewport.columns[0][0], CellValue::Number(0.0));
    assert_eq!(viewport.columns[1][0], CellValue::String("AA".to_string()));
    assert_eq!(viewport.columns[2][0], CellValue::Boolean(true));
}
