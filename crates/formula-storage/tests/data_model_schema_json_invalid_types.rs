use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions, Value};
use formula_dax::{DataModel, Table, Value as DaxValue};
use formula_storage::Storage;
use rusqlite::{params, Connection};
use tempfile::NamedTempFile;

#[test]
fn load_data_model_tolerates_invalid_schema_json() {
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
    let schema = vec![ColumnSchema {
        name: "Amount".to_string(),
        column_type: ColumnType::Number,
    }];
    let mut builder = ColumnarTableBuilder::new(schema, options);
    builder.append_row(&[Value::Number(10.0)]);
    builder.append_row(&[Value::Number(20.0)]);
    builder.append_row(&[Value::Number(30.0)]);
    let table = builder.finalize();

    let mut model = DataModel::new();
    model
        .add_table(Table::from_columnar("Fact", table))
        .expect("add table");
    storage
        .save_data_model(workbook.id, &model)
        .expect("save data model");
    drop(storage);

    // Corrupt the JSON schema blob for the table so deserialization fails. The loader should
    // still be able to infer page sizing from the stored chunks.
    let conn = Connection::open(path).expect("open raw db");
    let workbook_id_str = workbook.id.to_string();
    conn.execute(
        "UPDATE data_model_tables SET schema_json = ?1 WHERE workbook_id = ?2 AND name = ?3",
        params!["{", &workbook_id_str, "Fact"],
    )
    .expect("corrupt schema_json");
    drop(conn);

    let storage = Storage::open_path(path).expect("reopen storage");
    let loaded = storage.load_data_model(workbook.id).expect("load data model");
    let table = loaded.table("Fact").expect("loaded table");

    assert_eq!(table.value(2, "Amount"), Some(DaxValue::from(30.0)));
}

#[test]
fn load_data_model_tolerates_invalid_encoding_json_types() {
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
    let schema = vec![ColumnSchema {
        name: "Amount".to_string(),
        column_type: ColumnType::Number,
    }];
    let mut builder = ColumnarTableBuilder::new(schema, options);
    builder.append_row(&[Value::Number(10.0)]);
    builder.append_row(&[Value::Number(20.0)]);
    builder.append_row(&[Value::Number(30.0)]);
    let table = builder.finalize();

    let mut model = DataModel::new();
    model
        .add_table(Table::from_columnar("Fact", table))
        .expect("add table");
    storage
        .save_data_model(workbook.id, &model)
        .expect("save data model");
    drop(storage);

    // Corrupt the `encoding_json` column for the persisted column. This field is not required to
    // decode the chunks, so the loader should not fail solely due to an invalid SQLite type.
    let conn = Connection::open(path).expect("open raw db");
    let workbook_id_str = workbook.id.to_string();
    let table_id: i64 = conn
        .query_row(
            "SELECT id FROM data_model_tables WHERE workbook_id = ?1 AND name = ?2",
            params![&workbook_id_str, "Fact"],
            |r| r.get(0),
        )
        .expect("select table id");
    conn.execute(
        "UPDATE data_model_columns SET encoding_json = 0 WHERE table_id = ?1",
        params![table_id],
    )
    .expect("corrupt encoding_json type");
    drop(conn);

    let storage = Storage::open_path(path).expect("reopen storage");
    let loaded = storage.load_data_model(workbook.id).expect("load data model");
    let table = loaded.table("Fact").expect("loaded table");

    assert_eq!(table.value(2, "Amount"), Some(DaxValue::from(30.0)));
}

