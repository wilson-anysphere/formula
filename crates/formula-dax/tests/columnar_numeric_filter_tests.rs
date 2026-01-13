use formula_dax::{ColumnarTableBackend, TableBackend, Value};
use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions};

fn options() -> TableOptions {
    TableOptions {
        page_size_rows: 4,
        cache: PageCacheConfig { max_entries: 2 },
    }
}

#[test]
fn columnar_backend_numeric_filters_match_columnar_semantics() {
    let schema = vec![ColumnSchema {
        name: "n".to_owned(),
        column_type: ColumnType::Number,
    }];
    let mut builder = ColumnarTableBuilder::new(schema, options());

    let rows = [
        formula_columnar::Value::Number(0.0),
        formula_columnar::Value::Number(-0.0),
        formula_columnar::Value::Number(f64::NAN),
        formula_columnar::Value::Null,
        formula_columnar::Value::Number(1.0),
        formula_columnar::Value::Number(f64::NAN),
    ];
    for v in rows {
        builder.append_row(&[v]);
    }

    let table = builder.finalize();
    let backend = ColumnarTableBackend::new(table);

    // -0.0 and 0.0 are treated as equal, and nulls are excluded.
    assert_eq!(backend.filter_eq(0, &Value::from(-0.0)).unwrap(), vec![0, 1]);
    assert_eq!(backend.filter_eq(0, &Value::from(0.0)).unwrap(), vec![0, 1]);

    // All NaNs are grouped/equal together (consistent with `GROUP BY`).
    assert_eq!(
        backend.filter_eq(0, &Value::from(f64::NAN)).unwrap(),
        vec![2, 5]
    );

    assert_eq!(
        backend
            .filter_in(0, &[Value::from(0.0), Value::from(f64::NAN)])
            .unwrap(),
        vec![0, 1, 2, 5]
    );
}

#[test]
fn columnar_backend_numeric_filters_on_int_backed_columns_work() {
    let schema = vec![ColumnSchema {
        name: "dt".to_owned(),
        column_type: ColumnType::DateTime,
    }];
    let mut builder = ColumnarTableBuilder::new(schema, options());

    let rows = [
        formula_columnar::Value::DateTime(10),
        formula_columnar::Value::Null,
        formula_columnar::Value::DateTime(11),
        formula_columnar::Value::DateTime(10),
    ];
    for v in rows {
        builder.append_row(&[v]);
    }

    let table = builder.finalize();
    let backend = ColumnarTableBackend::new(table);

    assert_eq!(backend.filter_eq(0, &Value::from(10.0)).unwrap(), vec![0, 3]);
    assert_eq!(
        backend
            .filter_in(0, &[Value::from(10.0), Value::from(11.0)])
            .unwrap(),
        vec![0, 2, 3]
    );
}

