use formula_io::{open_workbook, save_workbook, Workbook};
use formula_model::{sanitize_sheet_name, CellValue};

#[test]
fn opens_csv_and_saves_as_xlsx() {
    let dir = tempfile::tempdir().expect("temp dir");
    let csv_path = dir.path().join("data.csv");
    std::fs::write(&csv_path, "col1,col2\n1,hello\n2,world\n").expect("write csv");

    let wb = open_workbook(&csv_path).expect("open csv workbook");
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
fn opens_windows1252_csv_and_saves_as_xlsx() {
    let dir = tempfile::tempdir().expect("temp dir");
    let csv_path = dir.path().join("data.csv");

    // "café" with Windows-1252 byte 0xE9 for "é" (invalid UTF-8).
    std::fs::write(&csv_path, b"col1,col2\n1,caf\xe9\n").expect("write csv");

    let wb = open_workbook(&csv_path).expect("open csv workbook");
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
        CellValue::String("café".to_string())
    );
}

#[test]
fn opens_utf8_bom_csv_schema_names_are_clean() {
    let dir = tempfile::tempdir().expect("temp dir");
    let csv_path = dir.path().join("bom.csv");

    std::fs::write(&csv_path, b"\xEF\xBB\xBFid,text\n1,hello\n").expect("write csv");

    let wb = open_workbook(&csv_path).expect("open csv workbook");
    let model = match wb {
        Workbook::Model(model) => model,
        other => panic!("expected Workbook::Model, got {other:?}"),
    };

    let sheet = model.sheet_by_name("bom").expect("sheet missing");
    let table = sheet.columnar_table().expect("expected columnar table");
    assert_eq!(table.schema()[0].name, "id");
}

#[test]
fn csv_import_sanitizes_sheet_name_from_file_stem() {
    let dir = tempfile::tempdir().expect("temp dir");

    // `[` and `]` are valid filename characters across platforms, but are invalid worksheet name
    // characters in Excel.
    let path = dir.path().join("bad[name].csv");

    std::fs::write(&path, "col1,col2\n1,hello\n2,world\n").expect("write csv");

    let wb = open_workbook(&path).expect("open csv workbook");
    let model = match wb {
        Workbook::Model(model) => model,
        other => panic!("expected Workbook::Model, got {other:?}"),
    };

    let expected = sanitize_sheet_name("bad[name]");
    assert_ne!(expected, "Sheet1", "expected sanitized name to not be the default");
    assert!(
        model.sheet_by_name("Sheet1").is_none(),
        "should not fall back to Sheet1 for an invalid but non-empty stem"
    );

    let sheet = model
        .sheet_by_name(&expected)
        .expect("expected worksheet name to be sanitized from file stem");

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
fn csv_import_invalid_sheet_name_falls_back_to_sheet1() {
    let dir = tempfile::tempdir().expect("temp dir");

    // Use a stem that becomes empty after Excel sheet-name sanitization.
    // `[` and `]` are invalid in sheet names but valid on common filesystems.
    let path = dir.path().join("[].csv");

    std::fs::write(&path, "col1\n1\n").expect("write csv");

    let wb = open_workbook(&path).expect("open csv workbook");
    let model = match wb {
        Workbook::Model(model) => model,
        other => panic!("expected Workbook::Model, got {other:?}"),
    };

    assert_eq!(sanitize_sheet_name("[]"), "Sheet1");
    let sheet = model.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
}

#[test]
fn opens_csv_with_wrong_extension() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("data.xlsx");

    // Note: the extension is intentionally wrong; content sniffing should still treat it as CSV.
    std::fs::write(&path, "col1,col2\n1,hello\n2,world\n").expect("write csv");

    let wb = open_workbook(&path).expect("open csv workbook");
    let model = match wb {
        Workbook::Model(model) => model,
        other => panic!("expected Workbook::Model, got {other:?}"),
    };

    let sheet_name = model.sheets[0].name.clone();
    let sheet = model.sheet_by_name(&sheet_name).expect("sheet missing");

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
fn opens_csv_with_xls_extension() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("data.xls");

    // Note: the extension is intentionally wrong; content sniffing should still treat it as CSV.
    std::fs::write(&path, "col1,col2\n1,hello\n2,world\n").expect("write csv");

    let wb = open_workbook(&path).expect("open csv workbook");
    let model = match wb {
        Workbook::Model(model) => model,
        other => panic!("expected Workbook::Model, got {other:?}"),
    };

    let sheet_name = model.sheets[0].name.clone();
    let sheet = model.sheet_by_name(&sheet_name).expect("sheet missing");

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
fn opens_csv_with_xlsb_extension() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("data.xlsb");

    // Note: the extension is intentionally wrong; content sniffing should still treat it as CSV.
    std::fs::write(&path, "col1,col2\n1,hello\n2,world\n").expect("write csv");

    let wb = open_workbook(&path).expect("open csv workbook");
    let model = match wb {
        Workbook::Model(model) => model,
        other => panic!("expected Workbook::Model, got {other:?}"),
    };

    let sheet_name = model.sheets[0].name.clone();
    let sheet = model.sheet_by_name(&sheet_name).expect("sheet missing");

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
fn opens_extensionless_csv() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("data");

    std::fs::write(&path, "col1,col2\n1,hello\n2,world\n").expect("write csv");

    let wb = open_workbook(&path).expect("open csv workbook");
    let model = match wb {
        Workbook::Model(model) => model,
        other => panic!("expected Workbook::Model, got {other:?}"),
    };

    let sheet_name = model.sheets[0].name.clone();
    let sheet = model.sheet_by_name(&sheet_name).expect("sheet missing");

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
fn opens_semicolon_delimited_csv_with_sniffed_delimiter() {
    let dir = tempfile::tempdir().expect("temp dir");
    let csv_path = dir.path().join("semi.csv");
    std::fs::write(&csv_path, "a;b\n1;2\n").expect("write csv");

    let wb = open_workbook(&csv_path).expect("open csv workbook");
    let model = match wb {
        Workbook::Model(model) => model,
        other => panic!("expected Workbook::Model, got {other:?}"),
    };

    let sheet = model.sheet_by_name("semi").expect("sheet missing");
    let table = sheet.columnar_table().expect("expected columnar table");
    assert_eq!(table.column_count(), 2);

    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
    assert_eq!(sheet.value_a1("B1").unwrap(), CellValue::Number(2.0));
}

#[test]
fn opens_tab_delimited_csv_with_sniffed_delimiter() {
    let dir = tempfile::tempdir().expect("temp dir");
    let csv_path = dir.path().join("tab.csv");
    std::fs::write(&csv_path, "a\tb\n1\t2\n").expect("write csv");

    let wb = open_workbook(&csv_path).expect("open csv workbook");
    let model = match wb {
        Workbook::Model(model) => model,
        other => panic!("expected Workbook::Model, got {other:?}"),
    };

    let sheet = model.sheet_by_name("tab").expect("sheet missing");
    let table = sheet.columnar_table().expect("expected columnar table");
    assert_eq!(table.column_count(), 2);

    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
    assert_eq!(sheet.value_a1("B1").unwrap(), CellValue::Number(2.0));
}

#[test]
fn opens_pipe_delimited_csv_with_sniffed_delimiter() {
    let dir = tempfile::tempdir().expect("temp dir");
    let csv_path = dir.path().join("pipe.csv");
    std::fs::write(&csv_path, "a|b\n1|2\n").expect("write csv");

    let wb = open_workbook(&csv_path).expect("open csv workbook");
    let model = match wb {
        Workbook::Model(model) => model,
        other => panic!("expected Workbook::Model, got {other:?}"),
    };

    let sheet = model.sheet_by_name("pipe").expect("sheet missing");
    let table = sheet.columnar_table().expect("expected columnar table");
    assert_eq!(table.column_count(), 2);

    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
    assert_eq!(sheet.value_a1("B1").unwrap(), CellValue::Number(2.0));
}

#[test]
fn opens_csv_with_excel_sep_directive() {
    let dir = tempfile::tempdir().expect("temp dir");
    let csv_path = dir.path().join("sep.csv");
    std::fs::write(&csv_path, "sep=;\na;b\n1;2\n").expect("write csv");

    let wb = open_workbook(&csv_path).expect("open csv workbook");
    let model = match wb {
        Workbook::Model(model) => model,
        other => panic!("expected Workbook::Model, got {other:?}"),
    };

    let sheet = model.sheet_by_name("sep").expect("sheet missing");
    let table = sheet.columnar_table().expect("expected columnar table");
    assert_eq!(table.column_count(), 2);

    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
    assert_eq!(sheet.value_a1("B1").unwrap(), CellValue::Number(2.0));
}
