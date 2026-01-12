use std::io::Cursor;

use formula_io::{save_workbook, Workbook};

#[test]
fn save_workbook_replaces_existing_file_atomically() {
    let mut model = formula_model::Workbook::new();
    model.add_sheet("Sheet1").expect("add sheet");

    let mut cursor = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&model, &mut cursor).expect("write workbook");
    let bytes = cursor.into_inner();
    let pkg = formula_xlsx::XlsxPackage::from_bytes(&bytes).expect("parse generated package");
    let expected = pkg.write_to_bytes().expect("write package bytes");

    let workbook = Workbook::Xlsx(pkg);

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("existing.xlsx");
    std::fs::write(&out_path, b"old-bytes").expect("seed existing file");

    save_workbook(&workbook, &out_path).expect("save workbook");
    let written = std::fs::read(&out_path).expect("read written bytes");
    assert_eq!(written, expected);
}

#[test]
fn save_workbook_creates_parent_directories() {
    let mut model = formula_model::Workbook::new();
    model.add_sheet("Sheet1").expect("add sheet");

    let mut cursor = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&model, &mut cursor).expect("write workbook");
    let bytes = cursor.into_inner();
    let pkg = formula_xlsx::XlsxPackage::from_bytes(&bytes).expect("parse generated package");
    let expected = pkg.write_to_bytes().expect("write package bytes");

    let workbook = Workbook::Xlsx(pkg);

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("nested/dir/out.xlsx");

    save_workbook(&workbook, &out_path).expect("save workbook");
    let written = std::fs::read(&out_path).expect("read written bytes");
    assert_eq!(written, expected);
}

