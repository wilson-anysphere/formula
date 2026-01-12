use formula_columnar::{ColumnSchema, ColumnType, ColumnarTableBuilder, PageCacheConfig, TableOptions, Value};
use formula_dax::{Cardinality, CrossFilterDirection, DataModel, Relationship, Table};
use formula_storage::Storage;
use rusqlite::{params, Connection};
use tempfile::NamedTempFile;

#[test]
fn load_data_model_schema_skips_relationship_rows_with_invalid_types() {
    let tmp = NamedTempFile::new().expect("tmpfile");
    let path = tmp.path();

    let storage = Storage::open_path(path).expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");

    // Create a minimal model with two tables so relationships table exists for this workbook id.
    let options = TableOptions {
        page_size_rows: 2,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let schema = vec![ColumnSchema {
        name: "Key".to_string(),
        column_type: ColumnType::Number,
    }];
    let mut builder = ColumnarTableBuilder::new(schema.clone(), options);
    builder.append_row(&[Value::Number(1.0)]);
    let t1 = builder.finalize();
    let mut builder = ColumnarTableBuilder::new(schema, options);
    builder.append_row(&[Value::Number(1.0)]);
    let t2 = builder.finalize();

    let mut model = DataModel::new();
    model
        .add_table(Table::from_columnar("T1", t1))
        .expect("add table");
    model
        .add_table(Table::from_columnar("T2", t2))
        .expect("add table");
    model
        .add_relationship(Relationship {
            name: "R1".to_string(),
            from_table: "T1".to_string(),
            from_column: "Key".to_string(),
            to_table: "T2".to_string(),
            to_column: "Key".to_string(),
            cardinality: Cardinality::OneToMany,
            cross_filter_direction: CrossFilterDirection::Single,
            is_active: true,
            enforce_referential_integrity: true,
        })
        .expect("add relationship");
    storage
        .save_data_model(workbook.id, &model)
        .expect("save data model");
    drop(storage);

    // Corrupt the relationships row with wrong SQLite types so decoding would fail.
    let conn = Connection::open(path).expect("open raw db");
    let workbook_id_str = workbook.id.to_string();
    conn.execute(
        "UPDATE data_model_relationships SET from_table = X'FF' WHERE workbook_id = ?1",
        params![&workbook_id_str],
    )
    .expect("corrupt relationship type");
    drop(conn);

    let storage = Storage::open_path(path).expect("reopen storage");
    let schema = storage
        .load_data_model_schema(workbook.id)
        .expect("load schema best-effort");

    assert!(
        schema.relationships.is_empty(),
        "expected invalid relationship rows to be skipped"
    );
}
