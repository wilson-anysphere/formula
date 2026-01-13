use formula_columnar::{
    ColumnAppendError, ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig,
    TableOptions, Value,
};
use std::sync::Arc;

#[test]
fn append_encoded_column_roundtrip() {
    let options = TableOptions {
        page_size_rows: 2,
        cache: PageCacheConfig { max_entries: 8 },
    };

    let base_schema = vec![
        ColumnSchema {
            name: "a".to_owned(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "b".to_owned(),
            column_type: ColumnType::String,
        },
    ];

    let mut base_builder = ColumnarTableBuilder::new(base_schema, options);
    for (a, b) in [
        (1.0, "A"),
        (2.0, "B"),
        (3.0, "C"),
        (4.0, "D"),
        (5.0, "E"),
    ] {
        base_builder.append_row(&[
            Value::Number(a),
            Value::String(Arc::<str>::from(b)),
        ]);
    }
    let base = base_builder.finalize();
    assert_eq!(base.row_count(), 5);
    assert_eq!(base.column_count(), 2);

    let extra_schema = vec![ColumnSchema {
        name: "c".to_owned(),
        column_type: ColumnType::Number,
    }];
    let mut extra_builder = ColumnarTableBuilder::new(extra_schema, options);
    for v in [10.0, 20.0, 30.0, 40.0, 50.0] {
        extra_builder.append_row(&[Value::Number(v)]);
    }
    let extra_table = extra_builder.finalize();
    let mut encoded = extra_table.into_encoded_columns();
    assert_eq!(encoded.len(), 1);
    let encoded_col = encoded.pop().unwrap();

    let appended = base.with_appended_encoded_column(encoded_col).unwrap();
    assert_eq!(appended.column_count(), 3);
    assert_eq!(appended.schema().len(), 3);

    // Existing columns remain readable.
    assert_eq!(appended.get_cell(0, 0), Value::Number(1.0));
    assert_eq!(appended.get_cell(4, 1), Value::String(Arc::<str>::from("E")));

    // Appended column reads back correctly.
    for (row, expected) in [10.0, 20.0, 30.0, 40.0, 50.0].into_iter().enumerate() {
        assert_eq!(appended.get_cell(row, 2), Value::Number(expected));
    }
}

#[test]
fn append_encoded_column_length_mismatch() {
    let options = TableOptions {
        page_size_rows: 2,
        cache: PageCacheConfig { max_entries: 8 },
    };

    let base_schema = vec![
        ColumnSchema {
            name: "a".to_owned(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "b".to_owned(),
            column_type: ColumnType::Number,
        },
    ];
    let mut base_builder = ColumnarTableBuilder::new(base_schema, options);
    for v in [1.0, 2.0, 3.0, 4.0, 5.0] {
        base_builder.append_row(&[Value::Number(v), Value::Number(v)]);
    }
    let base = base_builder.finalize();

    let extra_schema = vec![ColumnSchema {
        name: "c".to_owned(),
        column_type: ColumnType::Number,
    }];
    let mut extra_builder = ColumnarTableBuilder::new(extra_schema, options);
    for v in [10.0, 20.0, 30.0, 40.0] {
        extra_builder.append_row(&[Value::Number(v)]);
    }
    let extra_table = extra_builder.finalize();
    let encoded_col = extra_table.into_encoded_columns().into_iter().next().unwrap();

    let err = base.with_appended_encoded_column(encoded_col).unwrap_err();
    assert!(
        matches!(
            err,
            ColumnAppendError::LengthMismatch {
                expected: 5,
                actual: 4
            }
        ),
        "unexpected error: {err:?}"
    );
}

#[test]
fn append_encoded_column_duplicate_name() {
    let options = TableOptions {
        page_size_rows: 2,
        cache: PageCacheConfig { max_entries: 8 },
    };

    let base_schema = vec![
        ColumnSchema {
            name: "a".to_owned(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "b".to_owned(),
            column_type: ColumnType::Number,
        },
    ];
    let mut base_builder = ColumnarTableBuilder::new(base_schema, options);
    for v in [1.0, 2.0, 3.0, 4.0, 5.0] {
        base_builder.append_row(&[Value::Number(v), Value::Number(v)]);
    }
    let base = base_builder.finalize();

    // Build an encoded column that tries to reuse an existing column name.
    let dup_schema = vec![ColumnSchema {
        name: "a".to_owned(),
        column_type: ColumnType::Number,
    }];
    let mut dup_builder = ColumnarTableBuilder::new(dup_schema, options);
    for v in [10.0, 20.0, 30.0, 40.0, 50.0] {
        dup_builder.append_row(&[Value::Number(v)]);
    }
    let dup_table = dup_builder.finalize();
    let encoded_col = dup_table.into_encoded_columns().into_iter().next().unwrap();

    let err = base.with_appended_encoded_column(encoded_col).unwrap_err();
    assert!(
        matches!(err, ColumnAppendError::DuplicateColumn { ref name } if name == "a"),
        "unexpected error: {err:?}"
    );
}
