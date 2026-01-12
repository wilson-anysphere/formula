use formula_columnar::{
    ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions, Value,
};
use formula_dax::{DataModel, Table, Value as DaxValue};
use formula_storage::Storage;
use rusqlite::{params, Connection};
use std::sync::Arc;
use tempfile::NamedTempFile;

#[test]
fn load_data_model_handles_dictionary_indices_out_of_bounds() {
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

    // Corrupt the dictionary-encoded chunk indices so they reference an out-of-bounds dictionary
    // entry. The loader should not panic when reading values; it should degrade to blanks.
    let conn = Connection::open(path).expect("open raw db");
    let workbook_id_str = workbook.id.to_string();
    let table_id: i64 = conn
        .query_row(
            "SELECT id FROM data_model_tables WHERE workbook_id = ?1 AND name = ?2",
            params![&workbook_id_str, "Fact"],
            |r| r.get(0),
        )
        .expect("select table id");
    let category_column_id: i64 = conn
        .query_row(
            "SELECT id FROM data_model_columns WHERE table_id = ?1 AND name = ?2",
            params![table_id, "Category"],
            |r| r.get(0),
        )
        .expect("select column id");
    let chunk_id: i64 = conn
        .query_row(
            "SELECT id FROM data_model_chunks WHERE column_id = ?1 AND kind = 'dict' ORDER BY chunk_index LIMIT 1",
            params![category_column_id],
            |r| r.get(0),
        )
        .expect("select chunk id");

    let mut corrupt = Vec::new();
    corrupt.push(1); // CHUNK_BLOB_VERSION
    corrupt.extend_from_slice(&(2u32).to_le_bytes()); // len
    corrupt.push(2); // U32SequenceEncoding::Rle
    corrupt.extend_from_slice(&(1u32).to_le_bytes()); // run count
    corrupt.extend_from_slice(&(999u32).to_le_bytes()); // value (out of bounds)
    corrupt.extend_from_slice(&(2u32).to_le_bytes()); // ends
    corrupt.push(0); // no validity bitmap

    conn.execute(
        "UPDATE data_model_chunks SET data = ?1 WHERE id = ?2",
        params![corrupt, chunk_id],
    )
    .expect("corrupt chunk indices");
    drop(conn);

    let storage = Storage::open_path(path).expect("reopen storage");
    let loaded = storage.load_data_model(workbook.id).expect("load data model");
    let table = loaded.table("Fact").expect("loaded table");

    assert_eq!(table.value(0, "Category"), Some(DaxValue::Blank));
    assert_eq!(table.value(0, "Amount"), Some(DaxValue::from(10.0)));
}

