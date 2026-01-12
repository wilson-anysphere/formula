use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions, Value};
use formula_dax::{DataModel, Table};
use formula_storage::Storage;
use rusqlite::{params, Connection};
use std::sync::Arc;
use tempfile::NamedTempFile;

#[test]
fn stream_data_model_column_chunks_skips_rows_with_invalid_types() {
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
    builder.append_row(&[
        Value::String(Arc::<str>::from("C")),
        Value::Number(30.0),
    ]);
    let table = builder.finalize();

    let mut model = DataModel::new();
    model
        .add_table(Table::from_columnar("Fact", table))
        .expect("add table");
    storage
        .save_data_model(workbook.id, &model)
        .expect("save data model");
    drop(storage);

    // Corrupt one chunk row by storing a non-BLOB SQLite type in the `data` column.
    let conn = Connection::open(path).expect("open raw db");
    let workbook_id_str = workbook.id.to_string();
    let table_id: i64 = conn
        .query_row(
            "SELECT id FROM data_model_tables WHERE workbook_id = ?1 AND name = ?2",
            params![&workbook_id_str, "Fact"],
            |r| r.get(0),
        )
        .expect("select table id");
    let column_id: i64 = conn
        .query_row(
            "SELECT id FROM data_model_columns WHERE table_id = ?1 AND name = ?2",
            params![table_id, "Amount"],
            |r| r.get(0),
        )
        .expect("select column id");
    let chunk_count: i64 = conn
        .query_row(
            "SELECT COUNT(*) FROM data_model_chunks WHERE column_id = ?1",
            params![column_id],
            |r| r.get(0),
        )
        .expect("count chunks");
    assert!(
        chunk_count > 1,
        "expected multiple chunks so one corrupt chunk can be skipped"
    );
    let chunk_id: i64 = conn
        .query_row(
            "SELECT id FROM data_model_chunks WHERE column_id = ?1 ORDER BY chunk_index LIMIT 1",
            params![column_id],
            |r| r.get(0),
        )
        .expect("select chunk id");
    conn.execute(
        "UPDATE data_model_chunks SET data = 0 WHERE id = ?1",
        params![chunk_id],
    )
    .expect("corrupt chunk data type");
    drop(conn);

    let storage = Storage::open_path(path).expect("reopen storage");
    let mut seen = Vec::new();
    storage
        .stream_data_model_column_chunks(workbook.id, "Fact", "Amount", |chunk| {
            seen.push(chunk.chunk_index);
            Ok(())
        })
        .expect("stream chunks best-effort");

    assert!(
        !seen.is_empty(),
        "expected at least one valid chunk to be streamed"
    );
}
