use formula_columnar::{
    AggSpec, ColumnSchema, ColumnType, ColumnarTable, ColumnarTableBuilder, PageCacheConfig,
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
