use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions, Value};
use formula_dax::{DataModel, Table};
use formula_storage::Storage;
use rusqlite::{params, Connection};
use tempfile::NamedTempFile;

#[test]
fn load_data_model_infers_page_size_rows_when_schema_json_is_corrupt() {
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

    // Corrupt the persisted table schema to use an invalid `page_size_rows` value. The loader
    // should infer a usable page size from the stored chunks rather than constructing a table that
    // panics on access (division/modulo by zero).
    let conn = Connection::open(path).expect("open raw db");
    let workbook_id_str = workbook.id.to_string();
    let schema_json = r#"{"version":1,"page_size_rows":0,"cache_max_entries":4}"#;
    conn.execute(
        "UPDATE data_model_tables SET schema_json = ?1 WHERE workbook_id = ?2 AND name = ?3",
        params![schema_json, &workbook_id_str, "Fact"],
    )
    .expect("corrupt schema_json");
    drop(conn);

    let storage = Storage::open_path(path).expect("reopen storage");
    let loaded = storage.load_data_model(workbook.id).expect("load data model");
    let table = loaded.table("Fact").expect("loaded table");
    let columnar = table.columnar_table().expect("columnar table");

    assert_eq!(columnar.get_cell(2, 0), Value::Number(30.0));
}

