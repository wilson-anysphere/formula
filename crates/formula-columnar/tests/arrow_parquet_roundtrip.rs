#![cfg(feature = "arrow")]

use formula_columnar::arrow::{columnar_to_record_batch, record_batch_to_columnar};
use formula_columnar::parquet::{
    read_parquet_bytes_to_columnar, read_parquet_to_columnar, write_columnar_to_parquet,
    write_columnar_to_parquet_bytes,
};
use formula_columnar::{
    ColumnSchema, ColumnType, ColumnarTable, ColumnarTableBuilder, PageCacheConfig, TableOptions,
    Value,
};
use std::io::Cursor;
use std::sync::Arc;

fn make_table() -> ColumnarTable {
    let schema = vec![
        ColumnSchema {
            name: "num".to_owned(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "flag".to_owned(),
            column_type: ColumnType::Boolean,
        },
        ColumnSchema {
            name: "cat".to_owned(),
            column_type: ColumnType::String,
        },
        ColumnSchema {
            name: "ts".to_owned(),
            column_type: ColumnType::DateTime,
        },
        ColumnSchema {
            name: "money".to_owned(),
            column_type: ColumnType::Currency { scale: 2 },
        },
        ColumnSchema {
            name: "pct".to_owned(),
            column_type: ColumnType::Percentage { scale: 3 },
        },
    ];

    let options = TableOptions {
        page_size_rows: 4,
        cache: PageCacheConfig { max_entries: 4 },
    };
    let mut builder = ColumnarTableBuilder::new(schema, options);

    let rows: Vec<Vec<Value>> = vec![
        vec![
            Value::Number(1.0),
            Value::Boolean(true),
            Value::String(Arc::<str>::from("A")),
            Value::DateTime(100),
            Value::Currency(12_345),
            Value::Percentage(6_789),
        ],
        vec![
            Value::Null,
            Value::Boolean(false),
            Value::String(Arc::<str>::from("B")),
            Value::DateTime(101),
            Value::Null,
            Value::Percentage(10),
        ],
        vec![
            Value::Number(3.5),
            Value::Null,
            Value::Null,
            Value::DateTime(102),
            Value::Currency(0),
            Value::Null,
        ],
        vec![
            Value::Number(4.0),
            Value::Boolean(true),
            Value::String(Arc::<str>::from("A")),
            Value::Null,
            Value::Currency(999),
            Value::Percentage(1),
        ],
        vec![
            Value::Number(5.0),
            Value::Boolean(false),
            Value::String(Arc::<str>::from("C")),
            Value::DateTime(103),
            Value::Currency(-5),
            Value::Percentage(0),
        ],
    ];

    for row in rows {
        builder.append_row(&row);
    }

    builder.finalize()
}

fn assert_tables_equal(a: &ColumnarTable, b: &ColumnarTable) {
    assert_eq!(a.row_count(), b.row_count());
    assert_eq!(a.column_count(), b.column_count());
    assert_eq!(a.schema(), b.schema());

    for row in 0..a.row_count() {
        for col in 0..a.column_count() {
            assert_eq!(a.get_cell(row, col), b.get_cell(row, col), "r={row} c={col}");
        }
    }

    // String columns should remain dictionary encoded after round trips.
    let dict_a = a.dictionary(2).expect("string column dictionary");
    let dict_b = b.dictionary(2).expect("string column dictionary");
    let as_vec =
        |d: &Arc<Vec<Arc<str>>>| d.iter().map(|s| s.as_ref().to_owned()).collect::<Vec<_>>();
    assert_eq!(as_vec(&dict_a), vec!["A", "B", "C"]);
    assert_eq!(as_vec(&dict_b), vec!["A", "B", "C"]);
}

#[test]
fn arrow_ipc_roundtrip_preserves_schema_and_values() -> Result<(), Box<dyn std::error::Error>> {
    use arrow_ipc::reader::StreamReader;
    use arrow_ipc::writer::StreamWriter;

    let table = make_table();
    let batch = columnar_to_record_batch(&table)?;

    let mut bytes = Vec::new();
    {
        let mut writer = StreamWriter::try_new(&mut bytes, batch.schema().as_ref())?;
        writer.write(&batch)?;
        writer.finish()?;
    }

    let mut reader = StreamReader::try_new(Cursor::new(bytes), None)?;
    let read_batch = reader
        .next()
        .transpose()?
        .expect("expected one record batch");

    let table2 = record_batch_to_columnar(&read_batch)?;
    assert_tables_equal(&table, &table2);
    Ok(())
}

#[test]
fn parquet_roundtrip_bytes_preserves_schema_and_values() -> Result<(), Box<dyn std::error::Error>> {
    let table = make_table();
    let bytes = write_columnar_to_parquet_bytes(&table)?;
    let table2 = read_parquet_bytes_to_columnar(&bytes)?;
    assert_tables_equal(&table, &table2);
    Ok(())
}

#[test]
fn parquet_roundtrip_path_preserves_schema_and_values() -> Result<(), Box<dyn std::error::Error>> {
    use std::time::{SystemTime, UNIX_EPOCH};

    let table = make_table();

    let unique = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_nanos();
    let path = std::env::temp_dir().join(format!("formula-columnar-{unique}.parquet"));

    write_columnar_to_parquet(&table, &path)?;
    let table2 = read_parquet_to_columnar(&path)?;
    let _ = std::fs::remove_file(&path);

    assert_tables_equal(&table, &table2);
    Ok(())
}
