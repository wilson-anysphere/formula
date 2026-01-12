#![cfg(feature = "parquet")]

use std::path::PathBuf;
use std::sync::Arc;

use formula_columnar::{
    parquet::write_columnar_to_parquet_bytes, ColumnSchema, ColumnType, ColumnarTableBuilder,
    TableOptions, Value,
};
use formula_io::{open_workbook, save_workbook, Workbook};
use formula_model::{sanitize_sheet_name, CellValue};

fn parquet_fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../packages/data-io/test/fixtures")
        .join(rel)
}

#[test]
fn opens_parquet_and_saves_as_xlsx() {
    let schema = vec![
        ColumnSchema {
            name: "col1".to_string(),
            column_type: ColumnType::Number,
        },
        ColumnSchema {
            name: "col2".to_string(),
            column_type: ColumnType::String,
        },
    ];

    let mut builder = ColumnarTableBuilder::new(schema, TableOptions::default());
    builder.append_row(&[
        Value::Number(1.0),
        Value::String(Arc::<str>::from("hello")),
    ]);
    builder.append_row(&[
        Value::Number(2.0),
        Value::String(Arc::<str>::from("world")),
    ]);
    let table = builder.finalize();

    let bytes = write_columnar_to_parquet_bytes(&table).expect("write parquet bytes");

    let dir = tempfile::tempdir().expect("temp dir");
    let parquet_path = dir.path().join("data.parquet");
    std::fs::write(&parquet_path, bytes).expect("write parquet file");

    let wb = open_workbook(&parquet_path).expect("open parquet workbook");
    match wb {
        Workbook::Model(_) => {}
        other => panic!("expected Workbook::Model, got {other:?}"),
    }

    let out_path = dir.path().join("out.xlsx");
    save_workbook(&wb, &out_path).expect("save xlsx");

    let file = std::fs::File::open(&out_path).expect("open output xlsx");
    let model = formula_io::xlsx::read_workbook_from_reader(file).expect("read output xlsx");
    let sheet = model.sheets.first().expect("sheet");

    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
    assert_eq!(
        sheet.value_a1("B1").unwrap(),
        CellValue::String("hello".to_string())
    );
    assert_eq!(sheet.value_a1("A2").unwrap(), CellValue::Number(2.0));
    assert_eq!(
        sheet.value_a1("B2").unwrap(),
        CellValue::String("world".to_string())
    );
}

#[test]
fn parquet_fixture_can_export_to_xlsx() {
    let parquet_path = parquet_fixture_path("simple.parquet");
    let wb = open_workbook(&parquet_path).expect("open parquet workbook");

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("export.xlsx");
    save_workbook(&wb, &out_path).expect("save workbook as xlsx");

    let file = std::fs::File::open(&out_path).expect("open exported xlsx");
    let exported = formula_xlsx::read_workbook_from_reader(file).expect("read exported workbook");
    let sheet = exported
        .sheet_by_name("simple")
        .expect("expected worksheet name to match file stem");

    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
    assert_eq!(
        sheet.value_a1("B1").unwrap(),
        CellValue::String("Alice".to_string())
    );
    assert_eq!(sheet.value_a1("C2").unwrap(), CellValue::Boolean(false));
    assert_eq!(sheet.value_a1("D3").unwrap(), CellValue::Number(3.75));
}

#[test]
fn parquet_import_sanitizes_sheet_name_from_file_stem() {
    let parquet_path = parquet_fixture_path("simple.parquet");

    let dir = tempfile::tempdir().expect("temp dir");
    let bad_path = dir.path().join("bad[name].parquet");
    std::fs::copy(&parquet_path, &bad_path).expect("copy parquet fixture");

    let wb = open_workbook(&bad_path).expect("open parquet workbook");

    let out_path = dir.path().join("export.xlsx");
    save_workbook(&wb, &out_path).expect("save workbook as xlsx");

    let file = std::fs::File::open(&out_path).expect("open exported xlsx");
    let exported = formula_xlsx::read_workbook_from_reader(file).expect("read exported workbook");

    let expected = sanitize_sheet_name("bad[name]");
    assert_ne!(expected, "Sheet1", "expected sanitized name to not be the default");
    assert!(
        exported.sheet_by_name("Sheet1").is_none(),
        "should not fall back to Sheet1 for an invalid but non-empty stem"
    );
    exported
        .sheet_by_name(&expected)
        .expect("expected worksheet name to be sanitized from file stem");
}

#[test]
fn parquet_opens_with_wrong_extension_and_sanitizes_sheet_name() {
    let parquet_path = parquet_fixture_path("simple.parquet");

    let dir = tempfile::tempdir().expect("temp dir");
    // Note: the extension is intentionally wrong; content sniffing should still treat it as Parquet.
    let bad_path = dir.path().join("bad[name].xlsx");
    std::fs::copy(&parquet_path, &bad_path).expect("copy parquet fixture");

    let wb = open_workbook(&bad_path).expect("open parquet workbook");
    let model = match wb {
        Workbook::Model(model) => model,
        other => panic!("expected Workbook::Model, got {other:?}"),
    };

    let expected = sanitize_sheet_name("bad[name]");
    let sheet = model
        .sheet_by_name(&expected)
        .expect("expected worksheet name to be sanitized from file stem");

    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
    assert_eq!(
        sheet.value_a1("B1").unwrap(),
        CellValue::String("Alice".to_string())
    );
    assert_eq!(sheet.value_a1("C2").unwrap(), CellValue::Boolean(false));
    assert_eq!(sheet.value_a1("D3").unwrap(), CellValue::Number(3.75));
}

#[test]
fn parquet_import_invalid_sheet_name_falls_back_to_sheet1() {
    let parquet_path = parquet_fixture_path("simple.parquet");

    let dir = tempfile::tempdir().expect("temp dir");
    // Use a filename stem that becomes empty after sheet-name sanitization.
    // `[` and `]` are invalid in Excel sheet names but valid on common filesystems.
    let bad_path = dir.path().join("[].parquet");
    std::fs::copy(&parquet_path, &bad_path).expect("copy parquet fixture");

    let wb = open_workbook(&bad_path).expect("open parquet workbook");
    let model = match wb {
        Workbook::Model(model) => model,
        other => panic!("expected Workbook::Model, got {other:?}"),
    };

    assert_eq!(sanitize_sheet_name("[]"), "Sheet1");
    model
        .sheet_by_name("Sheet1")
        .expect("expected worksheet to fall back to Sheet1");
}
