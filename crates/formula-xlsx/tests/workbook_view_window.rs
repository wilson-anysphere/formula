use std::io::{Cursor, Read};

use formula_model::{Workbook, WorkbookWindow, WorkbookWindowState};
use formula_xlsx::XlsxDocument;
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
fn new_document_writes_workbook_view_window_metadata() {
    let mut workbook = Workbook::new();
    workbook.add_sheet("First").expect("add sheet");
    let sheet2 = workbook.add_sheet("Second").expect("add sheet");

    assert!(workbook.set_active_sheet(sheet2));
    workbook.view.window = Some(WorkbookWindow {
        x: Some(10),
        y: Some(20),
        width: Some(800),
        height: Some(600),
        state: Some(WorkbookWindowState::Maximized),
    });

    let doc = XlsxDocument::new(workbook);
    let bytes = doc.save_to_vec().expect("save xlsx");
    let workbook_xml = zip_part(&bytes, "xl/workbook.xml");
    let workbook_xml_str = std::str::from_utf8(&workbook_xml).expect("workbook.xml utf-8");

    let parsed = roxmltree::Document::parse(workbook_xml_str).expect("parse workbook.xml");
    let workbook_view = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "workbookView")
        .expect("expected <workbookView>");

    assert_eq!(workbook_view.attribute("activeTab"), Some("1"));
    assert_eq!(workbook_view.attribute("xWindow"), Some("10"));
    assert_eq!(workbook_view.attribute("yWindow"), Some("20"));
    assert_eq!(workbook_view.attribute("windowWidth"), Some("800"));
    assert_eq!(workbook_view.attribute("windowHeight"), Some("600"));
    assert_eq!(workbook_view.attribute("windowState"), Some("maximized"));
}

