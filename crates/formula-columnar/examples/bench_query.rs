use formula_columnar::{
    AggSpec, ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions, Value,
};
use std::sync::Arc;
use std::time::Instant;

fn main() {
    let rows: usize = std::env::var("ROWS")
        .ok()
        .and_then(|v| v.parse().ok())
        .unwrap_or(1_000_000);

    let options = TableOptions {
        page_size_rows: 65_536,
        cache: PageCacheConfig { max_entries: 32 },
    };

    // GROUP BY benchmark: sum by a dictionary-encoded string key.
    let categories: Vec<Arc<str>> = (0..1024)
        .map(|i| Arc::<str>::from(format!("C{i:04}")))
        .collect();

    let schema = vec![
        ColumnSchema {
            name: "category".to_owned(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "value".to_owned(),
            column_type: ColumnType::Number,
        },
    ];

    let start_build = Instant::now();
    let mut builder = ColumnarTableBuilder::new(schema, options);
    for i in 0..rows {
        builder.append_row(&[
            Value::String(categories[i % categories.len()].clone()),
            Value::Number((i % 100) as f64),
        ]);
    }
    let table = builder.finalize();
    let build_time = start_build.elapsed();

    let start_group_by = Instant::now();
    let gb = table
        .group_by(&[0], &[AggSpec::sum_f64(1)])
        .expect("group-by should succeed");
    let group_by_time = start_group_by.elapsed();
    println!("group_by groups: {}", gb.row_count());
    println!("group_by time: {:?}", group_by_time);
    println!("build time: {:?}", build_time);

    // JOIN benchmark: 1:1 inner join on an integer key.
    let join_schema = vec![ColumnSchema {
        name: "id".to_owned(),
        column_type: ColumnType::DateTime,
    }];
    let options = TableOptions {
        page_size_rows: 65_536,
        cache: PageCacheConfig { max_entries: 32 },
    };

    let mut left_builder = ColumnarTableBuilder::new(join_schema.clone(), options);
    let mut right_builder = ColumnarTableBuilder::new(join_schema, options);
    for i in 0..rows {
        let v = Value::DateTime(i as i64);
        left_builder.append_row(&[v.clone()]);
        right_builder.append_row(&[v]);
    }
    let left = left_builder.finalize();
    let right = right_builder.finalize();

    let start_join = Instant::now();
    let join = left
        .hash_join(&right, 0, 0)
        .expect("hash join should succeed");
    let join_time = start_join.elapsed();
    println!("join matches: {}", join.len());
    println!("join time: {:?}", join_time);
}

