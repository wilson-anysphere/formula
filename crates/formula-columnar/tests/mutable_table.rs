use formula_columnar::{
    ColumnSchema, ColumnType, ColumnarTable, ColumnarTableBuilder, EncodedColumn,
    MutableColumnarTable, PageCacheConfig, TableOptions, Value,
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

#[test]
fn delete_rows_rebuilds_and_shifts_indices() {
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
    for i in 0..8 {
        table.append_row(&[
            Value::Number(i as f64),
            Value::String(Arc::<str>::from(if i % 2 == 0 { "A" } else { "B" })),
        ]);
    }

    // Create an overlay update that should survive the delete via rebuild.
    assert!(table.update_cell(5, 0, Value::Number(100.0)));
    assert!(table.overlay_cell_count() > 0);

    let deleted = table.delete_rows(2, 4);
    assert_eq!(deleted, 2);
    assert_eq!(table.row_count(), 6);
    assert_eq!(table.overlay_cell_count(), 0);

    // Rows 0..2 preserved, row indices after deletion shift down.
    assert_eq!(table.get_cell(0, 0), Value::Number(0.0));
    assert_eq!(table.get_cell(1, 0), Value::Number(1.0));
    assert_eq!(table.get_cell(2, 0), Value::Number(4.0));
    assert_eq!(table.get_cell(3, 0), Value::Number(100.0)); // old row 5
    assert_eq!(table.get_cell(4, 0), Value::Number(6.0));
    assert_eq!(table.get_cell(5, 0), Value::Number(7.0));

    // Ensure we can still append after deletion without corrupting page boundaries.
    table.append_row(&[
        Value::Number(8.0),
        Value::String(Arc::<str>::from("C")),
    ]);
    table.append_row(&[
        Value::Number(9.0),
        Value::String(Arc::<str>::from("D")),
    ]);
    assert_eq!(table.row_count(), 8);
    assert_eq!(table.get_cell(6, 0), Value::Number(8.0));
    assert_eq!(table.get_cell(7, 0), Value::Number(9.0));
}

#[test]
fn distinct_count_survives_snapshot_and_append() {
    let schema = vec![ColumnSchema {
        name: "x".to_owned(),
        column_type: ColumnType::Number,
    }];
    let options = TableOptions {
        page_size_rows: 4,
        cache: PageCacheConfig { max_entries: 4 },
    };

    let mut builder = ColumnarTableBuilder::new(schema, options);
    for v in [1.0, 2.0, 3.0] {
        builder.append_row(&[Value::Number(v)]);
    }
    let table = builder.finalize();
    assert_eq!(table.scan().stats(0).unwrap().distinct_count, 3);

    let mut mutable = table.into_mutable();
    mutable.append_row(&[Value::Number(3.0)]);
    mutable.append_row(&[Value::Number(4.0)]);

    let stats = mutable.column_stats(0).unwrap();
    assert_eq!(stats.distinct_count, 4);
}

#[test]
fn distinct_count_after_from_encoded_and_append() {
    let schema = vec![ColumnSchema {
        name: "x".to_owned(),
        column_type: ColumnType::Number,
    }];
    let options = TableOptions {
        page_size_rows: 4,
        cache: PageCacheConfig { max_entries: 4 },
    };

    let mut builder = ColumnarTableBuilder::new(schema.clone(), options);
    for v in [1.0, 2.0, 3.0] {
        builder.append_row(&[Value::Number(v)]);
    }
    let table = builder.finalize();

    let encoded = vec![EncodedColumn {
        schema: schema[0].clone(),
        chunks: table.encoded_chunks(0).unwrap().to_vec(),
        stats: table.stats(0).unwrap().clone(),
        dictionary: table.dictionary(0),
    }];

    let restored = ColumnarTable::from_encoded(schema, encoded, table.row_count(), options);
    let mut mutable = restored.into_mutable();
    mutable.append_row(&[Value::Number(3.0)]);
    mutable.append_row(&[Value::Number(4.0)]);

    assert_eq!(mutable.column_stats(0).unwrap().distinct_count, 4);
}

#[test]
fn distinct_count_stats_canonicalizes_negative_zero_and_nans() {
    let schema = vec![ColumnSchema {
        name: "x".to_owned(),
        column_type: ColumnType::Number,
    }];
    let options = TableOptions {
        page_size_rows: 4,
        cache: PageCacheConfig { max_entries: 4 },
    };

    let mut builder = ColumnarTableBuilder::new(schema, options);
    builder.append_row(&[Value::Number(0.0)]);
    builder.append_row(&[Value::Number(-0.0)]);
    builder.append_row(&[Value::Number(f64::NAN)]);
    builder.append_row(&[Value::Number(f64::from_bits(0x7ff8_0000_0000_0001))]);
    builder.append_row(&[Value::Number(1.0)]);
    builder.append_row(&[Value::Null]);
    let table = builder.finalize();

    let stats = table.scan().stats(0).unwrap();
    assert_eq!(stats.null_count, 1);
    assert_eq!(stats.distinct_count, 3);

    // Ensure the mutable overlay keeps using the same canonicalization.
    let mut mutable = table.into_mutable();
    mutable.append_row(&[Value::Number(-0.0)]);
    mutable.append_row(&[Value::Number(f64::from_bits(0x7ff8_0000_0000_0010))]);
    assert_eq!(mutable.column_stats(0).unwrap().distinct_count, 3);
}

#[test]
fn append_after_into_mutable_from_partial_last_page() {
    let schema = vec![ColumnSchema {
        name: "x".to_owned(),
        column_type: ColumnType::Number,
    }];
    let options = TableOptions {
        page_size_rows: 4,
        cache: PageCacheConfig { max_entries: 4 },
    };

    let mut builder = ColumnarTableBuilder::new(schema.clone(), options);
    for v in 0..6 {
        builder.append_row(&[Value::Number(v as f64)]);
    }
    let table = builder.finalize();
    assert_eq!(table.row_count(), 6);

    // Immutable table has a partial last chunk (len=2). Ensure we can convert to mutable and
    // append without corrupting the row/page mapping.
    let mut mutable = table.into_mutable();
    mutable.append_row(&[Value::Number(6.0)]);
    mutable.append_row(&[Value::Number(7.0)]);

    assert_eq!(mutable.row_count(), 8);
    assert_eq!(mutable.get_cell(0, 0), Value::Number(0.0));
    assert_eq!(mutable.get_cell(5, 0), Value::Number(5.0));
    assert_eq!(mutable.get_cell(6, 0), Value::Number(6.0));
    assert_eq!(mutable.get_cell(7, 0), Value::Number(7.0));
}

#[test]
fn freeze_merges_overlays_across_flushed_and_current_pages() {
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
    for (x, cat) in [
        (0.0, "A"),
        (1.0, "B"),
        (2.0, "C"),
        (3.0, "D"),
        (4.0, "E"),
        (5.0, "F"),
    ] {
        table.append_row(&[
            Value::Number(x),
            Value::String(Arc::<str>::from(cat)),
        ]);
    }

    // Update a value in a flushed page (row 1) and in the current append buffer (row 5).
    assert!(table.update_cell(1, 0, Value::Number(100.0)));
    assert!(table.update_cell(1, 1, Value::String(Arc::<str>::from("Z"))));
    assert!(table.update_cell(5, 0, Value::Number(200.0)));
    assert!(table.update_cell(5, 1, Value::String(Arc::<str>::from("Y"))));

    let frozen = table.freeze();
    assert_eq!(frozen.row_count(), 6);
    assert_eq!(frozen.get_cell(1, 0), Value::Number(100.0));
    assert_eq!(frozen.get_cell(1, 1), Value::String(Arc::<str>::from("Z")));
    assert_eq!(frozen.get_cell(5, 0), Value::Number(200.0));
    assert_eq!(frozen.get_cell(5, 1), Value::String(Arc::<str>::from("Y")));
}

#[test]
fn append_after_compact_keeps_tail_buffer_readable() {
    let schema = vec![ColumnSchema {
        name: "x".to_owned(),
        column_type: ColumnType::Number,
    }];
    let options = TableOptions {
        page_size_rows: 4,
        cache: PageCacheConfig { max_entries: 4 },
    };

    let mut table = MutableColumnarTable::new(schema, options);
    for v in 0..6 {
        table.append_row(&[Value::Number(v as f64)]);
    }

    let snapshot = table.compact();
    assert_eq!(snapshot.row_count(), 6);
    assert_eq!(snapshot.get_cell(5, 0), Value::Number(5.0));

    // Append more after compact; new rows should be readable and should not corrupt earlier rows.
    table.append_row(&[Value::Number(6.0)]);
    table.append_row(&[Value::Number(7.0)]);
    assert_eq!(table.row_count(), 8);
    assert_eq!(table.get_cell(0, 0), Value::Number(0.0));
    assert_eq!(table.get_cell(5, 0), Value::Number(5.0));
    assert_eq!(table.get_cell(6, 0), Value::Number(6.0));
    assert_eq!(table.get_cell(7, 0), Value::Number(7.0));
}
