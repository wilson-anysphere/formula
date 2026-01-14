use formula_columnar::{
    AggSpec, BitVec, ColumnSchema, ColumnType, ColumnarTable, ColumnarTableBuilder, PageCacheConfig,
    TableOptions, Value,
};
use std::sync::Arc;

fn options() -> TableOptions {
    TableOptions {
        page_size_rows: 128,
        cache: PageCacheConfig { max_entries: 4 },
    }
}

fn build_table(schema: Vec<ColumnSchema>, rows: Vec<Vec<Value>>) -> ColumnarTable {
    let mut builder = ColumnarTableBuilder::new(schema, options());
    for row in rows {
        builder.append_row(&row);
    }
    builder.finalize()
}

#[test]
fn group_by_empty_table_is_empty() {
    let schema = vec![
        ColumnSchema {
            name: "k".to_owned(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "v".to_owned(),
            column_type: ColumnType::Number,
        },
    ];
    let table = ColumnarTableBuilder::new(schema, options()).finalize();

    let result = table
        .group_by(&[0], &[AggSpec::count_rows(), AggSpec::sum_f64(1)])
        .unwrap();

    assert_eq!(result.row_count(), 0);
    assert_eq!(result.column_count(), 3);

    let cols = result.to_values();
    assert_eq!(cols.len(), 3);
    assert!(cols.iter().all(|c| c.is_empty()));
}

#[test]
fn group_by_handles_null_keys_and_null_values() {
    let schema = vec![
        ColumnSchema {
            name: "k".to_owned(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "v".to_owned(),
            column_type: ColumnType::Number,
        },
    ];
    let rows = vec![
        vec![Value::String(Arc::<str>::from("A")), Value::Number(1.0)],
        vec![Value::Null, Value::Number(2.0)],
        vec![Value::Null, Value::Null],
        vec![Value::String(Arc::<str>::from("A")), Value::Number(3.0)],
        vec![Value::String(Arc::<str>::from("B")), Value::Null],
    ];
    let table = build_table(schema, rows);

    let result = table
        .group_by(
            &[0],
            &[
                AggSpec::count_rows().with_name("cnt_rows"),
                AggSpec::count_non_null(1).with_name("cnt_nn"),
                AggSpec::sum_f64(1),
                AggSpec::min(1),
                AggSpec::max(1),
            ],
        )
        .unwrap();

    assert_eq!(result.row_count(), 3);
    let cols = result.to_values();

    let mut lookup = std::collections::HashMap::<String, (Value, Value, Value, Value)>::new();
    for r in 0..result.row_count() {
        let key = &cols[0][r];
        let key_str = match key {
            Value::Null => "<null>".to_owned(),
            Value::String(s) => s.as_ref().to_owned(),
            other => format!("{other:?}"),
        };

        lookup.insert(
            key_str,
            (
                cols[1][r].clone(), // cnt_rows
                cols[2][r].clone(), // cnt_nn
                cols[3][r].clone(), // sum
                cols[4][r].clone(), // min
            ),
        );
        // max is cols[5], asserted below per key.
    }

    let a = lookup.get("A").unwrap();
    assert_eq!(&a.0, &Value::Number(2.0));
    assert_eq!(&a.1, &Value::Number(2.0));
    assert_eq!(&a.2, &Value::Number(4.0));
    assert_eq!(&a.3, &Value::Number(1.0));

    let idx_a = (0..result.row_count())
        .find(|&r| cols[0][r] == Value::String(Arc::<str>::from("A")))
        .unwrap();
    assert_eq!(cols[5][idx_a], Value::Number(3.0));

    let n = lookup.get("<null>").unwrap();
    assert_eq!(&n.0, &Value::Number(2.0));
    assert_eq!(&n.1, &Value::Number(1.0));
    assert_eq!(&n.2, &Value::Number(2.0));
    assert_eq!(&n.3, &Value::Number(2.0));

    let idx_null = (0..result.row_count())
        .find(|&r| cols[0][r] == Value::Null)
        .unwrap();
    assert_eq!(cols[5][idx_null], Value::Number(2.0));

    let b = lookup.get("B").unwrap();
    assert_eq!(&b.0, &Value::Number(1.0));
    assert_eq!(&b.1, &Value::Number(0.0));
    assert_eq!(&b.2, &Value::Null);
    assert_eq!(&b.3, &Value::Null);

    let idx_b = (0..result.row_count())
        .find(|&r| cols[0][r] == Value::String(Arc::<str>::from("B")))
        .unwrap();
    assert_eq!(cols[5][idx_b], Value::Null);
}

#[test]
fn group_by_high_cardinality_keys() {
    let n = 10_000;
    let schema = vec![ColumnSchema {
        name: "k".to_owned(),
        column_type: ColumnType::DateTime,
    }];
    let mut builder = ColumnarTableBuilder::new(schema, options());
    for i in 0..n {
        builder.append_row(&[Value::DateTime(i as i64)]);
    }
    let table = builder.finalize();

    let result = table.group_by(&[0], &[AggSpec::count_rows()]).unwrap();
    assert_eq!(result.row_count(), n);

    let cols = result.to_values();
    assert_eq!(cols[0][0], Value::DateTime(0));
    assert_eq!(cols[0][n - 1], Value::DateTime((n - 1) as i64));
    assert!(cols[1].iter().all(|v| v == &Value::Number(1.0)));
}

#[test]
fn group_by_skew_duplicates() {
    let schema = vec![
        ColumnSchema {
            name: "k".to_owned(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "v".to_owned(),
            column_type: ColumnType::Number,
        },
    ];

    let mut rows = Vec::new();
    for _ in 0..1_000 {
        rows.push(vec![Value::String(Arc::<str>::from("hot")), Value::Number(1.0)]);
    }
    for _ in 0..10 {
        rows.push(vec![Value::String(Arc::<str>::from("cold")), Value::Number(2.0)]);
    }
    let table = build_table(schema, rows);

    let result = table
        .group_by(&[0], &[AggSpec::count_rows(), AggSpec::sum_f64(1)])
        .unwrap();
    assert_eq!(result.row_count(), 2);
    let cols = result.to_values();

    let mut pairs: Vec<(Value, Value, Value)> = Vec::new();
    for r in 0..result.row_count() {
        pairs.push((cols[0][r].clone(), cols[1][r].clone(), cols[2][r].clone()));
    }
    pairs.sort_by(|a, b| format!("{:?}", a.0).cmp(&format!("{:?}", b.0)));

    // cold: 10 rows, sum 20
    assert_eq!(pairs[0].0, Value::String(Arc::<str>::from("cold")));
    assert_eq!(pairs[0].1, Value::Number(10.0));
    assert_eq!(pairs[0].2, Value::Number(20.0));

    // hot: 1000 rows, sum 1000
    assert_eq!(pairs[1].0, Value::String(Arc::<str>::from("hot")));
    assert_eq!(pairs[1].1, Value::Number(1000.0));
    assert_eq!(pairs[1].2, Value::Number(1000.0));
}

#[test]
fn group_by_rows_matches_group_by_on_selected_rows() {
    let schema = vec![
        ColumnSchema {
            name: "k".to_owned(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "v".to_owned(),
            column_type: ColumnType::Number,
        },
    ];
    let rows = vec![
        vec![Value::String(Arc::<str>::from("A")), Value::Number(1.0)],
        vec![Value::String(Arc::<str>::from("B")), Value::Number(2.0)],
        vec![Value::Null, Value::Number(3.0)],
        vec![Value::String(Arc::<str>::from("A")), Value::Null],
        vec![Value::String(Arc::<str>::from("C")), Value::Number(4.0)],
    ];
    let table = build_table(schema.clone(), rows.clone());

    let selected = vec![0usize, 2, 3, 4];
    let filtered_rows: Vec<Vec<Value>> = selected.iter().map(|&i| rows[i].clone()).collect();
    let filtered_table = build_table(schema, filtered_rows);

    let expected = filtered_table
        .group_by(&[0], &[AggSpec::count_rows(), AggSpec::sum_f64(1)])
        .unwrap()
        .to_values();

    let actual = table
        .group_by_rows(&[0], &[AggSpec::count_rows(), AggSpec::sum_f64(1)], &selected)
        .unwrap()
        .to_values();

    fn as_map(cols: &[Vec<Value>]) -> std::collections::HashMap<String, (Value, Value)> {
        let mut out = std::collections::HashMap::new();
        for row in 0..cols[0].len() {
            let key = match &cols[0][row] {
                Value::Null => "<null>".to_owned(),
                Value::String(s) => s.as_ref().to_owned(),
                other => format!("{other:?}"),
            };
            out.insert(key, (cols[1][row].clone(), cols[2][row].clone()));
        }
        out
    }

    assert_eq!(as_map(&expected), as_map(&actual));
}

#[test]
fn group_by_mask_matches_group_by_rows() {
    let schema = vec![
        ColumnSchema {
            name: "k".to_owned(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "v".to_owned(),
            column_type: ColumnType::Number,
        },
    ];
    let rows = vec![
        vec![Value::String(Arc::<str>::from("A")), Value::Number(1.0)],
        vec![Value::String(Arc::<str>::from("B")), Value::Number(2.0)],
        vec![Value::Null, Value::Number(3.0)],
        vec![Value::String(Arc::<str>::from("A")), Value::Null],
        vec![Value::String(Arc::<str>::from("C")), Value::Number(4.0)],
    ];
    let table = build_table(schema, rows);

    let selected = vec![0usize, 2, 3, 4];
    let mut mask = BitVec::with_len_all_false(table.row_count());
    for &row in &selected {
        mask.set(row, true);
    }

    let keys = [0usize];
    let aggs = [AggSpec::count_rows(), AggSpec::sum_f64(1)];

    let expected = table
        .group_by_rows(&keys, &aggs, &selected)
        .unwrap()
        .to_values();
    let actual = table.group_by_mask(&keys, &aggs, &mask).unwrap().to_values();

    fn as_map(cols: &[Vec<Value>]) -> std::collections::HashMap<String, (Value, Value)> {
        let mut out = std::collections::HashMap::new();
        for row in 0..cols[0].len() {
            let key = match &cols[0][row] {
                Value::Null => "<null>".to_owned(),
                Value::String(s) => s.as_ref().to_owned(),
                other => format!("{other:?}"),
            };
            out.insert(key, (cols[1][row].clone(), cols[2][row].clone()));
        }
        out
    }

    assert_eq!(as_map(&expected), as_map(&actual));
}

#[test]
fn group_by_avg_f64_ignores_nulls_and_nulls_when_no_values() {
    let schema = vec![
        ColumnSchema {
            name: "k".to_owned(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "v".to_owned(),
            column_type: ColumnType::Number,
        },
    ];
    let rows = vec![
        vec![Value::String(Arc::<str>::from("A")), Value::Number(1.0)],
        vec![Value::String(Arc::<str>::from("A")), Value::Null],
        vec![Value::String(Arc::<str>::from("A")), Value::Number(3.0)],
        vec![Value::String(Arc::<str>::from("B")), Value::Null],
        vec![Value::String(Arc::<str>::from("B")), Value::Null],
        vec![Value::Null, Value::Number(2.0)],
    ];
    let table = build_table(schema, rows);

    let result = table
        .group_by(&[0], &[AggSpec::avg_f64(1).with_name("avg")])
        .unwrap();
    assert_eq!(result.row_count(), 3);

    let cols = result.to_values();
    let mut lookup = std::collections::HashMap::<String, Value>::new();
    for r in 0..result.row_count() {
        let key_str = match &cols[0][r] {
            Value::Null => "<null>".to_owned(),
            Value::String(s) => s.as_ref().to_owned(),
            other => format!("{other:?}"),
        };
        lookup.insert(key_str, cols[1][r].clone());
    }

    assert_eq!(lookup.get("A"), Some(&Value::Number(2.0)));
    assert_eq!(lookup.get("B"), Some(&Value::Null));
    assert_eq!(lookup.get("<null>"), Some(&Value::Number(2.0)));
}

#[test]
fn group_by_distinct_count_ignores_nulls_and_outputs_zero() {
    let schema = vec![
        ColumnSchema {
            name: "k".to_owned(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "v".to_owned(),
            column_type: ColumnType::Number,
        },
    ];
    let rows = vec![
        vec![Value::String(Arc::<str>::from("A")), Value::Number(1.0)],
        vec![Value::String(Arc::<str>::from("A")), Value::Number(1.0)],
        vec![Value::String(Arc::<str>::from("A")), Value::Null],
        vec![Value::String(Arc::<str>::from("A")), Value::Number(2.0)],
        vec![Value::String(Arc::<str>::from("B")), Value::Null],
        vec![Value::String(Arc::<str>::from("B")), Value::Null],
        vec![Value::String(Arc::<str>::from("B")), Value::Number(3.0)],
        vec![Value::String(Arc::<str>::from("C")), Value::Null],
    ];
    let table = build_table(schema, rows);

    let result = table
        .group_by(&[0], &[AggSpec::distinct_count(1)])
        .unwrap();
    assert_eq!(result.row_count(), 3);

    let cols = result.to_values();
    let mut lookup = std::collections::HashMap::<String, Value>::new();
    for r in 0..result.row_count() {
        let key_str = match &cols[0][r] {
            Value::Null => "<null>".to_owned(),
            Value::String(s) => s.as_ref().to_owned(),
            other => format!("{other:?}"),
        };
        lookup.insert(key_str, cols[1][r].clone());
    }

    assert_eq!(lookup.get("A"), Some(&Value::Number(2.0)));
    assert_eq!(lookup.get("B"), Some(&Value::Number(1.0)));
    assert_eq!(lookup.get("C"), Some(&Value::Number(0.0)));
}

#[test]
fn group_by_distinct_count_strings_uses_dictionary_indices() {
    let schema = vec![
        ColumnSchema {
            name: "k".to_owned(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "s".to_owned(),
            column_type: ColumnType::String,
        },
    ];
    let rows = vec![
        vec![
            Value::String(Arc::<str>::from("G1")),
            Value::String(Arc::<str>::from("apple")),
        ],
        vec![
            Value::String(Arc::<str>::from("G1")),
            Value::String(Arc::<str>::from("banana")),
        ],
        vec![
            Value::String(Arc::<str>::from("G1")),
            Value::String(Arc::<str>::from("apple")),
        ],
        vec![Value::String(Arc::<str>::from("G1")), Value::Null],
        vec![
            Value::String(Arc::<str>::from("G2")),
            Value::String(Arc::<str>::from("banana")),
        ],
        vec![
            Value::String(Arc::<str>::from("G2")),
            Value::String(Arc::<str>::from("banana")),
        ],
        vec![Value::String(Arc::<str>::from("G2")), Value::Null],
        vec![Value::String(Arc::<str>::from("G3")), Value::Null],
    ];
    let table = build_table(schema, rows);

    let result = table
        .group_by(&[0], &[AggSpec::distinct_count(1)])
        .unwrap();
    assert_eq!(result.row_count(), 3);

    let cols = result.to_values();
    let mut lookup = std::collections::HashMap::<String, Value>::new();
    for r in 0..result.row_count() {
        let key_str = match &cols[0][r] {
            Value::Null => "<null>".to_owned(),
            Value::String(s) => s.as_ref().to_owned(),
            other => format!("{other:?}"),
        };
        lookup.insert(key_str, cols[1][r].clone());
    }

    assert_eq!(lookup.get("G1"), Some(&Value::Number(2.0)));
    assert_eq!(lookup.get("G2"), Some(&Value::Number(1.0)));
    assert_eq!(lookup.get("G3"), Some(&Value::Number(0.0)));
}

#[test]
fn group_by_distinct_count_number_canonicalizes_negative_zero_and_nans() {
    let schema = vec![
        ColumnSchema {
            name: "k".to_owned(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "v".to_owned(),
            column_type: ColumnType::Number,
        },
    ];
    let rows = vec![
        vec![Value::String(Arc::<str>::from("A")), Value::Number(0.0)],
        vec![Value::String(Arc::<str>::from("A")), Value::Number(-0.0)],
        vec![
            Value::String(Arc::<str>::from("A")),
            Value::Number(f64::from_bits(0x7ff8000000000001)),
        ],
        vec![
            Value::String(Arc::<str>::from("A")),
            Value::Number(f64::from_bits(0x7ff8000000000002)),
        ],
        vec![Value::String(Arc::<str>::from("A")), Value::Null],
    ];
    let table = build_table(schema, rows);

    let result = table
        .group_by(&[0], &[AggSpec::distinct_count(1)])
        .unwrap();
    assert_eq!(result.row_count(), 1);
    let cols = result.to_values();
    assert_eq!(cols[1][0], Value::Number(2.0));
}

#[test]
fn group_by_multiple_keys_and_multiple_new_aggs_smoke() {
    let schema = vec![
        ColumnSchema {
            name: "region".to_owned(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "active".to_owned(),
            column_type: ColumnType::Boolean,
        },
        ColumnSchema {
            name: "amount".to_owned(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "user".to_owned(),
            column_type: ColumnType::String,
        },
    ];
    let rows = vec![
        vec![
            Value::String(Arc::<str>::from("East")),
            Value::Boolean(true),
            Value::Number(10.0),
            Value::String(Arc::<str>::from("u1")),
        ],
        vec![
            Value::String(Arc::<str>::from("East")),
            Value::Boolean(true),
            Value::Number(20.0),
            Value::String(Arc::<str>::from("u2")),
        ],
        vec![
            Value::String(Arc::<str>::from("East")),
            Value::Boolean(true),
            Value::Null,
            Value::String(Arc::<str>::from("u2")),
        ],
        vec![
            Value::String(Arc::<str>::from("East")),
            Value::Boolean(false),
            Value::Number(5.0),
            Value::String(Arc::<str>::from("u1")),
        ],
        vec![
            Value::String(Arc::<str>::from("West")),
            Value::Boolean(true),
            Value::Null,
            Value::Null,
        ],
        vec![
            Value::String(Arc::<str>::from("West")),
            Value::Boolean(true),
            Value::Number(7.0),
            Value::String(Arc::<str>::from("u3")),
        ],
        vec![
            Value::String(Arc::<str>::from("West")),
            Value::Boolean(true),
            Value::Number(9.0),
            Value::String(Arc::<str>::from("u3")),
        ],
        vec![
            Value::Null,
            Value::Boolean(true),
            Value::Number(1.0),
            Value::String(Arc::<str>::from("u4")),
        ],
    ];
    let table = build_table(schema, rows);

    let result = table
        .group_by(
            &[0, 1],
            &[
                AggSpec::count_rows().with_name("cnt"),
                AggSpec::avg_f64(2).with_name("avg"),
                AggSpec::count_numbers(2).with_name("cnt_numbers"),
                AggSpec::distinct_count(3).with_name("dc_users"),
                AggSpec::var_p(2).with_name("varp"),
                AggSpec::std_dev_p(2).with_name("stddevp"),
            ],
        )
        .unwrap();

    assert_eq!(result.row_count(), 4);
    let cols = result.to_values();

    let mut lookup = std::collections::HashMap::<String, Vec<Value>>::new();
    for r in 0..result.row_count() {
        let region = match &cols[0][r] {
            Value::Null => "<null>".to_owned(),
            Value::String(s) => s.as_ref().to_owned(),
            other => format!("{other:?}"),
        };
        let active = match &cols[1][r] {
            Value::Boolean(b) => b.to_string(),
            Value::Null => "<null>".to_owned(),
            other => format!("{other:?}"),
        };
        let key = format!("{region}|{active}");
        lookup.insert(
            key,
            vec![
                cols[2][r].clone(), // cnt
                cols[3][r].clone(), // avg
                cols[4][r].clone(), // cnt_numbers
                cols[5][r].clone(), // dc_users
                cols[6][r].clone(), // varp
                cols[7][r].clone(), // stddevp
            ],
        );
    }

    assert_eq!(
        lookup.get("East|true"),
        Some(&vec![
            Value::Number(3.0),
            Value::Number(15.0),
            Value::Number(2.0),
            Value::Number(2.0),
            Value::Number(25.0),
            Value::Number(5.0),
        ])
    );
    assert_eq!(
        lookup.get("East|false"),
        Some(&vec![
            Value::Number(1.0),
            Value::Number(5.0),
            Value::Number(1.0),
            Value::Number(1.0),
            Value::Number(0.0),
            Value::Number(0.0),
        ])
    );
    assert_eq!(
        lookup.get("West|true"),
        Some(&vec![
            Value::Number(3.0),
            Value::Number(8.0),
            Value::Number(2.0),
            Value::Number(1.0),
            Value::Number(1.0),
            Value::Number(1.0),
        ])
    );
    assert_eq!(
        lookup.get("<null>|true"),
        Some(&vec![
            Value::Number(1.0),
            Value::Number(1.0),
            Value::Number(1.0),
            Value::Number(1.0),
            Value::Number(0.0),
            Value::Number(0.0),
        ])
    );
}

#[test]
fn group_by_var_and_stddev_sample_vs_population_semantics() {
    let schema = vec![
        ColumnSchema {
            name: "k".to_owned(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "v".to_owned(),
            column_type: ColumnType::Number,
        },
    ];
    let rows = vec![
        vec![Value::String(Arc::<str>::from("A")), Value::Number(1.0)],
        vec![Value::String(Arc::<str>::from("A")), Value::Number(3.0)],
        vec![Value::String(Arc::<str>::from("A")), Value::Null],
        vec![Value::String(Arc::<str>::from("B")), Value::Number(5.0)],
        vec![Value::String(Arc::<str>::from("C")), Value::Null],
    ];
    let table = build_table(schema, rows);

    let result = table
        .group_by(
            &[0],
            &[
                AggSpec::var(1).with_name("var_s"),
                AggSpec::std_dev(1).with_name("std_s"),
                AggSpec::var_p(1).with_name("var_p"),
                AggSpec::std_dev_p(1).with_name("std_p"),
            ],
        )
        .unwrap();
    assert_eq!(result.row_count(), 3);

    let cols = result.to_values();
    let mut lookup = std::collections::HashMap::<String, (Value, Value, Value, Value)>::new();
    for r in 0..result.row_count() {
        let key = match &cols[0][r] {
            Value::Null => "<null>".to_owned(),
            Value::String(s) => s.as_ref().to_owned(),
            other => format!("{other:?}"),
        };
        lookup.insert(
            key,
            (
                cols[1][r].clone(), // var_s
                cols[2][r].clone(), // std_s
                cols[3][r].clone(), // var_p
                cols[4][r].clone(), // std_p
            ),
        );
    }

    // A: values {1,3} -> mean 2
    // sample var = ((-1)^2 + (1)^2)/(2-1) = 2
    // pop var = ((-1)^2 + (1)^2)/2 = 1
    let a = lookup.get("A").unwrap();
    assert_eq!(a.0, Value::Number(2.0));
    assert_eq!(a.1, Value::Number((2.0f64).sqrt()));
    assert_eq!(a.2, Value::Number(1.0));
    assert_eq!(a.3, Value::Number(1.0));

    // B: single value -> sample var/stddev are NULL, population var/stddev are 0.
    let b = lookup.get("B").unwrap();
    assert_eq!(b.0, Value::Null);
    assert_eq!(b.1, Value::Null);
    assert_eq!(b.2, Value::Number(0.0));
    assert_eq!(b.3, Value::Number(0.0));

    // C: all null -> all outputs NULL.
    let c = lookup.get("C").unwrap();
    assert_eq!(c.0, Value::Null);
    assert_eq!(c.1, Value::Null);
    assert_eq!(c.2, Value::Null);
    assert_eq!(c.3, Value::Null);
}

#[test]
fn group_by_distinct_count_boolean_and_datetime_types() {
    // Boolean distinct count.
    let schema_bool = vec![
        ColumnSchema {
            name: "k".to_owned(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "b".to_owned(),
            column_type: ColumnType::Boolean,
        },
    ];
    let rows_bool = vec![
        vec![Value::String(Arc::<str>::from("A")), Value::Boolean(true)],
        vec![Value::String(Arc::<str>::from("A")), Value::Boolean(false)],
        vec![Value::String(Arc::<str>::from("A")), Value::Boolean(true)],
        vec![Value::String(Arc::<str>::from("A")), Value::Null],
        vec![Value::String(Arc::<str>::from("B")), Value::Null],
        vec![Value::String(Arc::<str>::from("C")), Value::Boolean(false)],
    ];
    let table_bool = build_table(schema_bool, rows_bool);
    let result_bool = table_bool
        .group_by(&[0], &[AggSpec::distinct_count(1)])
        .unwrap()
        .to_values();
    let mut lookup_bool = std::collections::HashMap::<String, Value>::new();
    for r in 0..result_bool[0].len() {
        let k = match &result_bool[0][r] {
            Value::String(s) => s.as_ref().to_owned(),
            Value::Null => "<null>".to_owned(),
            other => format!("{other:?}"),
        };
        lookup_bool.insert(k, result_bool[1][r].clone());
    }
    assert_eq!(lookup_bool.get("A"), Some(&Value::Number(2.0)));
    assert_eq!(lookup_bool.get("B"), Some(&Value::Number(0.0)));
    assert_eq!(lookup_bool.get("C"), Some(&Value::Number(1.0)));

    // DateTime (int-backed) distinct count.
    let schema_dt = vec![
        ColumnSchema {
            name: "k".to_owned(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "t".to_owned(),
            column_type: ColumnType::DateTime,
        },
    ];
    let rows_dt = vec![
        vec![Value::String(Arc::<str>::from("A")), Value::DateTime(1)],
        vec![Value::String(Arc::<str>::from("A")), Value::DateTime(1)],
        vec![Value::String(Arc::<str>::from("A")), Value::DateTime(2)],
        vec![Value::String(Arc::<str>::from("A")), Value::Null],
        vec![Value::String(Arc::<str>::from("B")), Value::Null],
        vec![Value::String(Arc::<str>::from("C")), Value::DateTime(-7)],
    ];
    let table_dt = build_table(schema_dt, rows_dt);
    let result_dt = table_dt
        .group_by(&[0], &[AggSpec::distinct_count(1)])
        .unwrap()
        .to_values();
    let mut lookup_dt = std::collections::HashMap::<String, Value>::new();
    for r in 0..result_dt[0].len() {
        let k = match &result_dt[0][r] {
            Value::String(s) => s.as_ref().to_owned(),
            Value::Null => "<null>".to_owned(),
            other => format!("{other:?}"),
        };
        lookup_dt.insert(k, result_dt[1][r].clone());
    }
    assert_eq!(lookup_dt.get("A"), Some(&Value::Number(2.0)));
    assert_eq!(lookup_dt.get("B"), Some(&Value::Number(0.0)));
    assert_eq!(lookup_dt.get("C"), Some(&Value::Number(1.0)));
}

#[test]
fn group_by_distinct_count_works_when_counting_a_key_column() {
    // When the agg input column is also a GROUP BY key column, the engine should reuse the key
    // scalar values (no extra decoding cursor). Since the key is constant within the group,
    // DISTINCTCOUNT should be 1 for non-null keys and 0 for null keys.
    let schema = vec![
        ColumnSchema {
            name: "k".to_owned(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "v".to_owned(),
            column_type: ColumnType::Number,
        },
    ];
    let rows = vec![
        vec![Value::String(Arc::<str>::from("A")), Value::Number(1.0)],
        vec![Value::String(Arc::<str>::from("A")), Value::Number(1.0)],
        vec![Value::String(Arc::<str>::from("A")), Value::Number(2.0)],
        vec![Value::String(Arc::<str>::from("A")), Value::Null],
        vec![Value::String(Arc::<str>::from("B")), Value::Null],
        vec![Value::String(Arc::<str>::from("B")), Value::Number(3.0)],
        vec![Value::Null, Value::Null],
        vec![Value::Null, Value::Number(4.0)],
    ];
    let table = build_table(schema, rows);

    let result = table
        .group_by(&[0, 1], &[AggSpec::distinct_count(1)])
        .unwrap();
    let cols = result.to_values();

    let mut lookup = std::collections::HashMap::<String, Value>::new();
    for r in 0..result.row_count() {
        let k = match &cols[0][r] {
            Value::Null => "<null>".to_owned(),
            Value::String(s) => s.as_ref().to_owned(),
            other => format!("{other:?}"),
        };
        let v = match &cols[1][r] {
            Value::Null => "<null>".to_owned(),
            Value::Number(n) => n.to_string(),
            other => format!("{other:?}"),
        };
        lookup.insert(format!("{k}|{v}"), cols[2][r].clone());
    }

    assert_eq!(lookup.get("A|1"), Some(&Value::Number(1.0)));
    assert_eq!(lookup.get("A|2"), Some(&Value::Number(1.0)));
    assert_eq!(lookup.get("A|<null>"), Some(&Value::Number(0.0)));
    assert_eq!(lookup.get("B|3"), Some(&Value::Number(1.0)));
    assert_eq!(lookup.get("B|<null>"), Some(&Value::Number(0.0)));
    assert_eq!(lookup.get("<null>|4"), Some(&Value::Number(1.0)));
    assert_eq!(lookup.get("<null>|<null>"), Some(&Value::Number(0.0)));
}

#[test]
fn hash_join_handles_duplicate_keys() {
    let schema = vec![ColumnSchema {
        name: "k".to_owned(),
        column_type: ColumnType::DateTime,
    }];
    let left = build_table(
        schema.clone(),
        vec![
            vec![Value::DateTime(1)],
            vec![Value::DateTime(1)],
            vec![Value::DateTime(2)],
        ],
    );
    let right = build_table(
        schema,
        vec![
            vec![Value::DateTime(1)],
            vec![Value::DateTime(1)],
            vec![Value::DateTime(1)],
            vec![Value::DateTime(3)],
        ],
    );

    let join = left.hash_join(&right, 0, 0).unwrap();
    assert_eq!(join.len(), 6);

    let mut pairs: Vec<(usize, usize)> = join
        .left_indices
        .into_iter()
        .zip(join.right_indices.into_iter())
        .collect();
    pairs.sort();

    assert_eq!(
        pairs,
        vec![(0, 0), (0, 1), (0, 2), (1, 0), (1, 1), (1, 2)]
    );
}

#[test]
fn hash_join_ignores_null_keys() {
    let schema = vec![ColumnSchema {
        name: "k".to_owned(),
        column_type: ColumnType::String,
    }];
    let left = build_table(
        schema.clone(),
        vec![
            vec![Value::String(Arc::<str>::from("A"))],
            vec![Value::Null],
            vec![Value::String(Arc::<str>::from("B"))],
        ],
    );
    let right = build_table(
        schema,
        vec![
            vec![Value::Null],
            vec![Value::String(Arc::<str>::from("A"))],
            vec![Value::String(Arc::<str>::from("B"))],
        ],
    );

    let join = left.hash_join(&right, 0, 0).unwrap();
    let mut pairs: Vec<(usize, usize)> = join
        .left_indices
        .into_iter()
        .zip(join.right_indices.into_iter())
        .collect();
    pairs.sort();
    assert_eq!(pairs, vec![(0, 1), (2, 2)]);
}

#[test]
fn hash_join_string_works_with_different_dictionaries() {
    let schema = vec![ColumnSchema {
        name: "k".to_owned(),
        column_type: ColumnType::String,
    }];
    let left = build_table(
        schema.clone(),
        vec![
            vec![Value::String(Arc::<str>::from("A"))],
            vec![Value::String(Arc::<str>::from("B"))],
            vec![Value::String(Arc::<str>::from("A"))],
        ],
    );
    // Insert in different order so the dictionaries differ.
    let right = build_table(
        schema,
        vec![
            vec![Value::String(Arc::<str>::from("B"))],
            vec![Value::String(Arc::<str>::from("A"))],
            vec![Value::String(Arc::<str>::from("A"))],
        ],
    );

    let join = left.hash_join(&right, 0, 0).unwrap();
    assert_eq!(join.len(), 5);

    let mut pairs: Vec<(usize, usize)> = join
        .left_indices
        .into_iter()
        .zip(join.right_indices.into_iter())
        .collect();
    pairs.sort();

    assert_eq!(
        pairs,
        vec![(0, 1), (0, 2), (1, 0), (2, 1), (2, 2)]
    );
}
