use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions, Value};
use formula_dax::{DataModel, Table};
use formula_storage::Storage;
use rusqlite::{params, Connection};
use tempfile::NamedTempFile;

#[test]
fn load_data_model_clamps_row_count_when_persisted_value_exceeds_chunk_rows() {
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

    // Corrupt the table row_count to exceed the number of values encoded in the chunks. Without
    // clamping, columnar access can panic by indexing past the end of the last chunk.
    let conn = Connection::open(path).expect("open raw db");
    let workbook_id_str = workbook.id.to_string();
    conn.execute(
        "UPDATE data_model_tables SET row_count = 4 WHERE workbook_id = ?1 AND name = ?2",
        params![&workbook_id_str, "Fact"],
    )
    .expect("corrupt row_count");
    drop(conn);

    let storage = Storage::open_path(path).expect("reopen storage");
    let loaded = storage.load_data_model(workbook.id).expect("load data model");
    let table = loaded.table("Fact").expect("loaded table");
    let columnar = table.columnar_table().expect("columnar table");

    assert_eq!(table.row_count(), 3);
    assert_eq!(columnar.get_cell(2, 0), Value::Number(30.0));
}

