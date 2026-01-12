use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions, Value};
use formula_dax::{DataModel, Table};
use formula_storage::Storage;
use rusqlite::{params, Connection};
use std::sync::Arc;
use tempfile::NamedTempFile;

#[test]
fn data_model_loads_best_effort_when_persisted_rows_are_corrupt() {
    let tmp = NamedTempFile::new().expect("tmpfile");
    let path = tmp.path();

    let storage = Storage::open_path(path).expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");

    let options = TableOptions {
        page_size_rows: 2,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let schema = vec![
        ColumnSchema {
            name: "Category".to_string(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "Amount".to_string(),
            column_type: ColumnType::Number,
        },
    ];
    let mut builder = ColumnarTableBuilder::new(schema, options);
    builder.append_row(&[
        Value::String(Arc::<str>::from("A")),
        Value::Number(10.0),
    ]);
    builder.append_row(&[
        Value::String(Arc::<str>::from("B")),
        Value::Number(20.0),
    ]);
    let table = builder.finalize();

    let mut model = DataModel::new();
    model
        .add_table(Table::from_columnar("Fact", table))
        .expect("add table");
    model
        .add_measure("Total", "SUM(Fact[Amount])")
        .expect("add measure");
    storage
        .save_data_model(workbook.id, &model)
        .expect("save data model");
    drop(storage);

    // Corrupt the stored column type JSON so schema/model loads can't deserialize it.
    let conn = Connection::open(path).expect("open raw db");
    conn.execute(
        "UPDATE data_model_columns SET column_type = '{' WHERE table_id IN (SELECT id FROM data_model_tables WHERE workbook_id = ?1)",
        params![workbook.id.to_string()],
    )
    .expect("corrupt column type");
    drop(conn);

    let storage = Storage::open_path(path).expect("reopen storage");
    let schema = storage
        .load_data_model_schema(workbook.id)
        .expect("schema-only load should be best-effort");
    assert_eq!(schema.tables.len(), 1);
    assert!(
        schema.tables[0].columns.is_empty(),
        "expected corrupt column type to be skipped"
    );

    let loaded = storage
        .load_data_model(workbook.id)
        .expect("full load should be best-effort");
    assert_eq!(
        loaded.tables().count(),
        0,
        "expected corrupt table to be skipped during full load"
    );
}

