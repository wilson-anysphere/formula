use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions, Value};
use formula_dax::{DataModel, Table};
use formula_storage::Storage;
use rusqlite::{params, Connection};
use tempfile::NamedTempFile;

#[test]
fn load_data_model_schema_tolerates_invalid_row_count_types() {
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

    // Corrupt the `row_count` column with a non-integer SQLite type.
    let conn = Connection::open(path).expect("open raw db");
    let workbook_id_str = workbook.id.to_string();
    conn.execute(
        "UPDATE data_model_tables SET row_count = 'bad' WHERE workbook_id = ?1 AND name = ?2",
        params![&workbook_id_str, "Fact"],
    )
    .expect("corrupt row_count type");
    drop(conn);

    let storage = Storage::open_path(path).expect("reopen storage");
    let schema = storage
        .load_data_model_schema(workbook.id)
        .expect("load schema best-effort");

    let fact = schema
        .tables
        .iter()
        .find(|t| t.name == "Fact")
        .expect("fact table present");
    assert_eq!(fact.row_count, 0);
    assert_eq!(fact.columns.len(), 1);
    assert_eq!(fact.columns[0].name, "Amount");
    assert_eq!(fact.columns[0].column_type, ColumnType::Number);
}

