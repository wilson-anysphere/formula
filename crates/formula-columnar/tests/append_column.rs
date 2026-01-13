use formula_columnar::{
    ColumnAppendError, ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions,
    Value,
};
use std::sync::Arc;

#[test]
fn append_numeric_column_happy_path() {
    let schema = vec![ColumnSchema {
        name: "a".to_owned(),
        column_type: ColumnType::Number,
    }];
    let options = TableOptions {
        page_size_rows: 4,
        cache: PageCacheConfig { max_entries: 8 },
    };

    let mut builder = ColumnarTableBuilder::new(schema, options);
    for v in [1.0, 2.0, 3.0] {
        builder.append_row(&[Value::Number(v)]);
    }
    let table = builder.finalize();
    assert_eq!(table.row_count(), 3);
    assert_eq!(table.column_count(), 1);

    let appended_schema = ColumnSchema {
        name: "b".to_owned(),
        column_type: ColumnType::Number,
    };
    let values = vec![Value::Number(10.0), Value::Number(20.0), Value::Number(30.0)];
    let table = table
        .with_appended_column(appended_schema, values)
        .expect("append column");

    assert_eq!(table.column_count(), 2);
    assert_eq!(table.get_cell(0, 1), Value::Number(10.0));
    assert_eq!(table.get_cell(1, 1), Value::Number(20.0));
    assert_eq!(table.get_cell(2, 1), Value::Number(30.0));
}

#[test]
fn append_string_column_is_dictionary_encoded() {
    let schema = vec![ColumnSchema {
        name: "id".to_owned(),
        column_type: ColumnType::Number,
    }];
    let options = TableOptions {
        page_size_rows: 4,
        cache: PageCacheConfig { max_entries: 8 },
    };

    let mut builder = ColumnarTableBuilder::new(schema, options);
    for v in [1.0, 2.0, 3.0, 4.0] {
        builder.append_row(&[Value::Number(v)]);
    }
    let table = builder.finalize();

    let appended_schema = ColumnSchema {
        name: "cat".to_owned(),
        column_type: ColumnType::String,
    };
    let values = vec![
        Value::String(Arc::<str>::from("A")),
        Value::String(Arc::<str>::from("B")),
        Value::String(Arc::<str>::from("A")),
        Value::Null,
    ];
    let table = table
        .with_appended_column(appended_schema, values)
        .expect("append column");

    let new_col = 1;
    let dict = table.dictionary(new_col).expect("dictionary encoded");
    assert!(!dict.is_empty(), "dictionary should be populated");

    assert_eq!(table.get_cell(0, new_col), Value::String(Arc::<str>::from("A")));
    assert_eq!(table.get_cell(1, new_col), Value::String(Arc::<str>::from("B")));
    assert_eq!(table.get_cell(2, new_col), Value::String(Arc::<str>::from("A")));
    assert_eq!(table.get_cell(3, new_col), Value::Null);
}

#[test]
fn append_column_rejects_duplicate_name() {
    let schema = vec![ColumnSchema {
        name: "a".to_owned(),
        column_type: ColumnType::Number,
    }];
    let options = TableOptions {
        page_size_rows: 4,
        cache: PageCacheConfig { max_entries: 8 },
    };

    let mut builder = ColumnarTableBuilder::new(schema, options);
    builder.append_row(&[Value::Number(1.0)]);
    builder.append_row(&[Value::Number(2.0)]);
    let table = builder.finalize();

    let err = table
        .with_appended_column(
            ColumnSchema {
                name: "a".to_owned(),
                column_type: ColumnType::Number,
            },
            vec![Value::Number(3.0), Value::Number(4.0)],
        )
        .expect_err("duplicate column should error");

    assert_eq!(
        err,
        ColumnAppendError::DuplicateColumn {
            name: "a".to_owned()
        }
    );
}

#[test]
fn append_column_rejects_length_mismatch() {
    let schema = vec![ColumnSchema {
        name: "a".to_owned(),
        column_type: ColumnType::Number,
    }];
    let options = TableOptions {
        page_size_rows: 4,
        cache: PageCacheConfig { max_entries: 8 },
    };

    let mut builder = ColumnarTableBuilder::new(schema, options);
    builder.append_row(&[Value::Number(1.0)]);
    builder.append_row(&[Value::Number(2.0)]);
    builder.append_row(&[Value::Number(3.0)]);
    let table = builder.finalize();

    let err = table
        .with_appended_column(
            ColumnSchema {
                name: "b".to_owned(),
                column_type: ColumnType::Number,
            },
            vec![Value::Number(10.0), Value::Number(20.0)],
        )
        .expect_err("length mismatch should error");

    assert_eq!(
        err,
        ColumnAppendError::LengthMismatch {
            expected: 3,
            actual: 2
        }
    );
}

#[test]
fn append_column_handles_partial_last_page() {
    let schema = vec![ColumnSchema {
        name: "a".to_owned(),
        column_type: ColumnType::Number,
    }];
    let options = TableOptions {
        page_size_rows: 4,
        cache: PageCacheConfig { max_entries: 8 },
    };

    // 6 rows with a 4-row page size => last page is partial.
    let mut builder = ColumnarTableBuilder::new(schema, options);
    for v in 0..6 {
        builder.append_row(&[Value::Number(v as f64)]);
    }
    let table = builder.finalize();
    assert_eq!(table.row_count(), 6);

    let appended_schema = ColumnSchema {
        name: "b".to_owned(),
        column_type: ColumnType::Number,
    };
    let values: Vec<Value> = (100..106).map(|v| Value::Number(v as f64)).collect();
    let table = table
        .with_appended_column(appended_schema, values)
        .expect("append column");

    let new_col = 1;
    assert_eq!(table.get_cell(3, new_col), Value::Number(103.0));
    assert_eq!(table.get_cell(4, new_col), Value::Number(104.0));
    assert_eq!(table.get_cell(5, new_col), Value::Number(105.0));
}

