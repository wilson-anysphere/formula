use std::io::{Cursor, Read};

use formula_model::{Workbook, WorkbookWindow, WorkbookWindowState};
use zip::ZipArchive;

fn zip_part(bytes: &[u8], name: &str) -> String {
    let cursor = Cursor::new(bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut out = String::new();
    file.read_to_string(&mut out).expect("read part");
    out
}

#[test]
fn write_workbook_emits_workbook_view_and_protection_elements() -> Result<(), Box<dyn std::error::Error>>
{
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1")?;

    workbook.workbook_protection.lock_structure = true;
    workbook.workbook_protection.lock_windows = true;
    workbook.workbook_protection.password_hash = Some(0x1A2B);

    workbook.view.window = Some(WorkbookWindow {
        x: Some(100),
        y: Some(200),
        width: Some(300),
        height: Some(400),
        state: Some(WorkbookWindowState::Maximized),
    });

    {
        let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
        sheet.sheet_protection.enabled = true;
        sheet.sheet_protection.password_hash = Some(0x00FF);
        sheet.sheet_protection.format_cells = true;
        sheet.sheet_protection.select_locked_cells = false;
        sheet.sheet_protection.edit_objects = true;
        sheet.sheet_protection.edit_scenarios = false;
    }

    let mut buf = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&workbook, &mut buf)?;
    let bytes = buf.into_inner();

    let workbook_xml = zip_part(&bytes, "xl/workbook.xml");
    let workbook_doc = roxmltree::Document::parse(&workbook_xml)?;
    let wb_prot = workbook_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "workbookProtection")
        .expect("expected <workbookProtection>");
    assert_eq!(wb_prot.attribute("lockStructure"), Some("1"));
    assert_eq!(wb_prot.attribute("lockWindows"), Some("1"));
    assert_eq!(wb_prot.attribute("workbookPassword"), Some("1A2B"));

    let wb_view = workbook_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "workbookView")
        .expect("expected <workbookView>");
    assert_eq!(wb_view.attribute("xWindow"), Some("100"));
    assert_eq!(wb_view.attribute("yWindow"), Some("200"));
    assert_eq!(wb_view.attribute("windowWidth"), Some("300"));
    assert_eq!(wb_view.attribute("windowHeight"), Some("400"));
    assert_eq!(wb_view.attribute("windowState"), Some("maximized"));

    let sheet_xml = zip_part(&bytes, "xl/worksheets/sheet1.xml");
    let sheet_doc = roxmltree::Document::parse(&sheet_xml)?;
    let sheet_prot = sheet_doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "sheetProtection")
        .expect("expected <sheetProtection>");
    assert_eq!(sheet_prot.attribute("sheet"), Some("1"));
    assert_eq!(sheet_prot.attribute("password"), Some("00FF"));
    assert_eq!(sheet_prot.attribute("formatCells"), Some("1"));
    assert_eq!(sheet_prot.attribute("selectLockedCells"), Some("0"));
    assert_eq!(sheet_prot.attribute("objects"), Some("0"));
    assert_eq!(sheet_prot.attribute("scenarios"), Some("1"));

    Ok(())
}
