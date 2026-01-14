use std::io::Read;

use formula_model::Workbook;
use formula_xlsx::{write_workbook_to_writer, XlsxDocument};
use zip::ZipArchive;

fn sheet1_xml(bytes: &[u8]) -> String {
    let cursor = std::io::Cursor::new(bytes);
    let mut archive = ZipArchive::new(cursor).expect("zip open");
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("sheet1.xml")
        .read_to_string(&mut sheet_xml)
        .expect("read sheet xml");
    sheet_xml
}

#[test]
fn xlsx_document_does_not_emit_negative_zero_in_sheet_xml() {
    // `XlsxDocument` uses the patch-based writer path; ensure it never emits `-0` for common
    // sizing attrs when the model contains `-0.0` (which can happen via arithmetic or user input).
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");
    {
        let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
        sheet.default_col_width = Some(-0.0);
        sheet.default_row_height = Some(-0.0);
        sheet.set_col_width(0, Some(-0.0));
        sheet.set_row_height(0, Some(-0.0));
        // Split panes use `xSplit`/`ySplit` attributes; ensure we don't emit `-0` there either.
        sheet.view.pane.x_split = Some(-0.0);
        sheet.view.pane.y_split = Some(-0.0);
    }

    let bytes = XlsxDocument::new(workbook).save_to_vec().expect("save");
    let sheet_xml = sheet1_xml(&bytes);

    assert!(
        !sheet_xml.contains("defaultColWidth=\"-0\""),
        "unexpected -0 in defaultColWidth: {sheet_xml}"
    );
    assert!(
        !sheet_xml.contains("defaultRowHeight=\"-0\""),
        "unexpected -0 in defaultRowHeight: {sheet_xml}"
    );
    assert!(
        !sheet_xml.contains("width=\"-0\""),
        "unexpected -0 in col width: {sheet_xml}"
    );
    assert!(
        !sheet_xml.contains("ht=\"-0\""),
        "unexpected -0 in row height: {sheet_xml}"
    );
    assert!(
        !sheet_xml.contains("xSplit=\"-0\""),
        "unexpected -0 in pane xSplit: {sheet_xml}"
    );
    assert!(
        !sheet_xml.contains("ySplit=\"-0\""),
        "unexpected -0 in pane ySplit: {sheet_xml}"
    );
}

#[test]
fn workbook_writer_does_not_emit_negative_zero_in_sheet_xml() {
    // `write_workbook_to_writer` uses the regeneration writer path; ensure it also avoids `-0`.
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");
    {
        let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
        sheet.default_col_width = Some(-0.0);
        sheet.default_row_height = Some(-0.0);
        sheet.set_col_width(0, Some(-0.0));
        sheet.set_row_height(0, Some(-0.0));
    }

    let mut cursor = std::io::Cursor::new(Vec::new());
    write_workbook_to_writer(&workbook, &mut cursor).expect("write xlsx");
    let bytes = cursor.into_inner();

    let sheet_xml = sheet1_xml(&bytes);
    assert!(
        !sheet_xml.contains("\"-0\""),
        "unexpected -0 in sheet XML: {sheet_xml}"
    );
}
