use std::io::{Cursor, Read};

use formula_model::{CellRef, CellValue, Workbook};
use formula_xlsx::{CellMeta, XlsxDocument};
use zip::ZipArchive;

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

#[test]
fn writer_emits_vm_cm_on_new_document() {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");
    workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        .set_value(
            CellRef::from_a1("A1").expect("valid A1"),
            CellValue::String("MSFT".to_string()),
        );

    let mut doc = XlsxDocument::new(workbook);
    doc.xlsx_meta_mut().cell_meta.insert(
        (sheet_id, CellRef::from_a1("A1").expect("valid A1")),
        CellMeta {
            vm: Some("1".to_string()),
            cm: Some("2".to_string()),
            ..Default::default()
        },
    );

    let saved = doc.save_to_vec().expect("save xlsx");

    let sheet_xml = zip_part(&saved, "xl/worksheets/sheet1.xml");
    let sheet_xml_str = std::str::from_utf8(&sheet_xml).expect("sheet1.xml utf-8");
    let parsed = roxmltree::Document::parse(sheet_xml_str).expect("parse sheet1.xml");

    let cell = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("expected A1 cell");

    assert_eq!(
        cell.attribute("vm"),
        Some("1"),
        "expected vm attribute to be written, got: {sheet_xml_str}"
    );
    assert_eq!(
        cell.attribute("cm"),
        Some("2"),
        "expected cm attribute to be written, got: {sheet_xml_str}"
    );
}
