use formula_columnar::{
    ColumnSchema, ColumnType, MutableColumnarTable, PageCacheConfig, TableOptions, Value,
};
use std::sync::Arc;

#[test]
fn append_rows_updates_reads_and_stats() {
    let schema = vec![
        ColumnSchema {
            name: "x".to_owned(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "cat".to_owned(),
            column_type: ColumnType::String,
        },
    ];
    let options = TableOptions {
        page_size_rows: 4,
        cache: PageCacheConfig { max_entries: 8 },
    };

    let mut table = MutableColumnarTable::new(schema, options);
    table.append_row(&[
        Value::Number(1.0),
        Value::String(Arc::<str>::from("A")),
    ]);
    table.append_row(&[
        Value::Number(2.0),
        Value::String(Arc::<str>::from("B")),
    ]);

    table.append_rows(&vec![
        vec![
            Value::Number(3.0),
            Value::String(Arc::<str>::from("A")),
        ],
        vec![
            Value::Number(4.0),
            Value::String(Arc::<str>::from("C")),
        ],
        vec![
            Value::Number(5.0),
            Value::String(Arc::<str>::from("D")),
        ],
    ]);

    assert_eq!(table.row_count(), 5);
    assert_eq!(table.get_cell(0, 0), Value::Number(1.0));
    assert_eq!(table.get_cell(4, 1), Value::String(Arc::<str>::from("D")));

    let dict = table.dictionary(1).expect("string column dictionary");
    let dict_vals: Vec<&str> = dict.iter().map(|s| s.as_ref()).collect();
    assert_eq!(dict_vals, vec!["A", "B", "C", "D"]);

    let stats_num = table.column_stats(0).unwrap();
    assert_eq!(stats_num.null_count, 0);
    assert_eq!(stats_num.distinct_count, 5);
    assert_eq!(stats_num.min, Some(Value::Number(1.0)));
    assert_eq!(stats_num.max, Some(Value::Number(5.0)));
    assert_eq!(stats_num.sum, Some(15.0));

    let stats_str = table.column_stats(1).unwrap();
    assert_eq!(stats_str.null_count, 0);
    assert_eq!(stats_str.distinct_count, 4);
    assert_eq!(stats_str.min, Some(Value::String(Arc::<str>::from("A"))));
    assert_eq!(stats_str.max, Some(Value::String(Arc::<str>::from("D"))));
    assert_eq!(stats_str.avg_length, Some(1.0));
}

#[test]
fn updates_are_visible_and_recompute_extremes() {
    let schema = vec![
        ColumnSchema {
            name: "x".to_owned(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "cat".to_owned(),
            column_type: ColumnType::String,
        },
    ];
    let options = TableOptions {
        page_size_rows: 8,
        cache: PageCacheConfig { max_entries: 4 },
    };

    let mut table = MutableColumnarTable::new(schema, options);
    for (x, cat) in [(1.0, "B"), (2.0, "A"), (3.0, "C")] {
        table.append_row(&[
            Value::Number(x),
            Value::String(Arc::<str>::from(cat)),
        ]);
    }

    assert_eq!(table.column_stats(0).unwrap().min, Some(Value::Number(1.0)));
    assert_eq!(
        table.column_stats(1).unwrap().min,
        Some(Value::String(Arc::<str>::from("A")))
    );

    // Overwrite the numeric minimum to force a min recompute.
    assert!(table.update_cell(0, 0, Value::Number(10.0)));
    assert_eq!(table.get_cell(0, 0), Value::Number(10.0));
    assert_eq!(table.column_stats(0).unwrap().min, Some(Value::Number(2.0)));

    // Update a small range (2 rows x 2 cols) in row-major order.
    let updated = table.update_range(
        1,
        3,
        0,
        2,
        &[
            Value::Number(20.0),
            Value::String(Arc::<str>::from("Z")),
            Value::Number(30.0),
            Value::String(Arc::<str>::from("Y")),
        ],
    );
    assert_eq!(updated, 4);
    assert_eq!(table.get_cell(1, 0), Value::Number(20.0));
    assert_eq!(table.get_cell(1, 1), Value::String(Arc::<str>::from("Z")));
    assert_eq!(table.get_cell(2, 0), Value::Number(30.0));
    assert_eq!(table.get_cell(2, 1), Value::String(Arc::<str>::from("Y")));
}

#[test]
fn compact_clears_overlays_and_matches_reads() {
    let schema = vec![
        ColumnSchema {
            name: "x".to_owned(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "cat".to_owned(),
            column_type: ColumnType::String,
        },
    ];
    let options = TableOptions {
        page_size_rows: 4,
        cache: PageCacheConfig { max_entries: 4 },
    };

    let mut table = MutableColumnarTable::new(schema, options);
    for (x, cat) in [(1.0, "A"), (2.0, "B"), (3.0, "C"), (4.0, "D")] {
        table.append_row(&[
            Value::Number(x),
            Value::String(Arc::<str>::from(cat)),
        ]);
    }

    assert!(table.update_cell(0, 0, Value::Number(100.0)));
    assert!(table.update_cell(3, 1, Value::String(Arc::<str>::from("Z"))));
    assert!(table.overlay_cell_count() > 0);

    let before = table.get_range(0, table.row_count(), 0, table.column_count());
    let frozen = table.compact();
    assert_eq!(table.overlay_cell_count(), 0);

    for r in 0..before.rows() {
        for c in 0..before.cols() {
            let expected = before.get(r, c).cloned().unwrap_or(Value::Null);
            assert_eq!(table.get_cell(r, c), expected);
            assert_eq!(frozen.get_cell(r, c), expected);
        }
    }
}

