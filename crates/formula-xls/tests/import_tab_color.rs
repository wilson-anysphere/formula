use std::io::{Cursor, Write};

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

fn export_xlsx_and_read_sheet_tab_color(
    workbook: &formula_model::Workbook,
    sheet_name: &str,
) -> Option<formula_model::TabColor> {
    let mut out = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(workbook, &mut out).expect("write xlsx");
    let bytes = out.into_inner();

    let sheets = formula_xlsx::worksheet_parts_from_reader(Cursor::new(bytes.as_slice()))
        .expect("worksheet parts");
    let sheet = sheets
        .iter()
        .find(|s| s.name == sheet_name)
        .unwrap_or_else(|| panic!("missing worksheet part for `{sheet_name}`"));

    let sheet_xml = formula_xlsx::read_part_from_reader(Cursor::new(bytes.as_slice()), &sheet.worksheet_part)
        .expect("read worksheet part")
        .expect("missing worksheet part");
    let sheet_xml = String::from_utf8(sheet_xml).expect("sheet xml utf-8");
    formula_xlsx::parse_sheet_tab_color(&sheet_xml)
        .expect("parse sheet tab color")
}

#[test]
fn imports_sheet_tab_color_from_sheetext() {
    let bytes = xls_fixture_builder::build_tab_color_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("TabColor")
        .expect("TabColor missing");

    let color = sheet.tab_color.as_ref().expect("tab_color missing");
    assert_eq!(color.rgb.as_deref(), Some("FF112233"));
    assert!(color.indexed.is_none(), "expected rgb-only tab color");

    // Ensure `.xls -> .xlsx` export preserves the tab color.
    let exported = export_xlsx_and_read_sheet_tab_color(&result.workbook, "TabColor")
        .expect("exported tab_color missing");
    assert_eq!(exported.rgb.as_deref(), Some("FF112233"));
    assert!(exported.indexed.is_none(), "expected rgb-only exported tab color");
}

#[test]
fn imports_sheet_tab_color_indexed_resolved_via_palette() {
    let bytes = xls_fixture_builder::build_tab_color_indexed_palette_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("TabColorIndexed")
        .expect("TabColorIndexed missing");

    let color = sheet.tab_color.as_ref().expect("tab_color missing");
    assert_eq!(color.rgb.as_deref(), Some("FF112233"));
    assert!(
        color.indexed.is_none(),
        "expected indexed tab color to be resolved to rgb"
    );

    // Ensure `.xls -> .xlsx` export preserves the resolved RGB tab color.
    let exported = export_xlsx_and_read_sheet_tab_color(&result.workbook, "TabColorIndexed")
        .expect("exported tab_color missing");
    assert_eq!(exported.rgb.as_deref(), Some("FF112233"));
    assert!(exported.indexed.is_none(), "expected rgb-only exported tab color");
}
