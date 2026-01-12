use std::io::{Cursor, Read};

use formula_model::Workbook;
use formula_xlsx::{write_workbook_to_writer_with_kind, WorkbookKind};
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
fn simple_writer_emits_active_tab_without_window_metadata() {
    let mut workbook = Workbook::new();
    workbook.add_sheet("First").expect("add sheet");
    let second = workbook.add_sheet("Second").expect("add sheet");
    assert!(workbook.set_active_sheet(second));
    workbook.view.window = None;

    let mut cursor = Cursor::new(Vec::new());
    write_workbook_to_writer_with_kind(&workbook, &mut cursor, WorkbookKind::Workbook)
        .expect("write workbook");
    let bytes = cursor.into_inner();

    let workbook_xml = zip_part(&bytes, "xl/workbook.xml");
    let workbook_xml_str = std::str::from_utf8(&workbook_xml).expect("workbook.xml utf-8");
    let parsed = roxmltree::Document::parse(workbook_xml_str).expect("parse workbook.xml");

    let workbook_view = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "workbookView")
        .expect("expected <workbookView>");

    assert_eq!(workbook_view.attribute("activeTab"), Some("1"));
}

