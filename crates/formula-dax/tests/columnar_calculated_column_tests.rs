use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions};
use formula_dax::{DataModel, DaxError, Table, Value};

#[test]
fn add_calculated_column_streams_into_columnar_table() {
    let schema = vec![ColumnSchema {
        name: "X".to_string(),
        column_type: ColumnType::Number,
    }];
    let options = TableOptions {
        page_size_rows: 16,
        cache: PageCacheConfig { max_entries: 4 },
    };

    let mut builder = ColumnarTableBuilder::new(schema, options);
    for i in 0..128 {
        builder.append_row(&[formula_columnar::Value::Number(i as f64)]);
    }

    let mut model = DataModel::new();
    model
        .add_table(Table::from_columnar("T", builder.finalize()))
        .unwrap();

    // Includes leading blanks to force "type inference while streaming" logic.
    model
        .add_calculated_column("T", "Y", "IF([X] < 5, BLANK(), [X] * 2)")
        .unwrap();

    let table = model.table("T").unwrap();
    assert_eq!(table.row_count(), 128);

    for i in 0..5 {
        assert_eq!(table.value(i, "Y").unwrap(), Value::Blank);
    }
    assert_eq!(table.value(5, "Y").unwrap(), 10.0.into());
    assert_eq!(table.value(10, "Y").unwrap(), 20.0.into());

    let columnar = table.columnar_table().unwrap();
    let y_schema = columnar.schema().iter().find(|c| c.name == "Y").unwrap();
    assert_eq!(y_schema.column_type, ColumnType::Number);
}

#[test]
fn columnar_calculated_column_type_mismatch_does_not_mutate_table() {
    let schema = vec![ColumnSchema {
        name: "X".to_string(),
        column_type: ColumnType::Number,
    }];
    let options = TableOptions {
        page_size_rows: 8,
        cache: PageCacheConfig { max_entries: 2 },
    };

    let mut builder = ColumnarTableBuilder::new(schema, options);
    for i in 0..32 {
        builder.append_row(&[formula_columnar::Value::Number(i as f64)]);
    }

    let mut model = DataModel::new();
    model
        .add_table(Table::from_columnar("T", builder.finalize()))
        .unwrap();

    let err = model
        .add_calculated_column("T", "Y", "IF([X] = 0, 1, \"A\")")
        .unwrap_err();
    assert!(matches!(err, DaxError::Type(_)));

    let table = model.table("T").unwrap();
    assert!(!table.columns().iter().any(|c| c == "Y"));
    assert_eq!(table.row_count(), 32);
    assert_eq!(table.value(0, "X").unwrap(), 0.0.into());
}
