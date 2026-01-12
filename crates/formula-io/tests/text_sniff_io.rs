use formula_io::{open_workbook, Error, Workbook};
use formula_model::CellValue;

#[test]
fn opens_csv_named_xlsx_via_content_sniffing() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("data.xlsx");
    std::fs::write(&path, "col1,col2\n1,hello\n2,world\n").expect("write csv bytes");

    let wb = open_workbook(&path).expect("open workbook");
    let model = match wb {
        Workbook::Model(model) => model,
        other => panic!("expected Workbook::Model, got {other:?}"),
    };

    let sheet = model.sheet_by_name("data").expect("data sheet missing");
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
fn opens_extensionless_csv_via_content_sniffing() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("data");
    std::fs::write(&path, "col1,col2\n1,hello\n2,world\n").expect("write csv bytes");

    let wb = open_workbook(&path).expect("open workbook");
    let model = match wb {
        Workbook::Model(model) => model,
        other => panic!("expected Workbook::Model, got {other:?}"),
    };

    let sheet = model.sheet_by_name("data").expect("data sheet missing");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
}

#[test]
fn opens_utf16le_tab_delimited_text_via_content_sniffing() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("data.txt");

    // UTF-16LE with BOM, matching Excel's "Unicode Text" export.
    let tsv = "col1\tcol2\n1\thello\n2\tworld\n";
    let mut bytes = vec![0xFF, 0xFE];
    for unit in tsv.encode_utf16() {
        bytes.extend_from_slice(&unit.to_le_bytes());
    }
    std::fs::write(&path, &bytes).expect("write utf16 tsv bytes");

    let wb = open_workbook(&path).expect("open workbook");
    let model = match wb {
        Workbook::Model(model) => model,
        other => panic!("expected Workbook::Model, got {other:?}"),
    };

    let sheet = model.sheet_by_name("data").expect("data sheet missing");
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
fn does_not_classify_binary_data_as_csv() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("binary");
    std::fs::write(&path, b"\x00\x01\x02\x03\x04\x05\n,\x06\x07").expect("write binary bytes");

    let err = open_workbook(&path).expect_err("expected open_workbook to fail");
    assert!(
        matches!(err, Error::UnsupportedExtension { .. }),
        "expected UnsupportedExtension, got {err:?}"
    );
}
