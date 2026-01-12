use std::io::Write;

use formula_model::{CellRef, Range};

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_workbook_and_sheet_view_state_from_biff() {
    let bytes = xls_fixture_builder::build_view_state_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet1 = result
        .workbook
        .sheet_by_name("Sheet1")
        .expect("Sheet1 missing");
    let sheet2 = result
        .workbook
        .sheet_by_name("Sheet2")
        .expect("Sheet2 missing");

    // Workbook WINDOW1.activeTab (itabCur) selects the second sheet.
    assert_eq!(result.workbook.view.active_sheet_id, Some(sheet2.id));
    assert_ne!(sheet1.id, sheet2.id);
    // The fixture's WINDOW1 does not set window geometry; ensure we don't persist a meaningless
    // 0x0 Normal window.
    assert!(result.workbook.view.window.is_none());

    // SCL zoom = 200% => 2.0
    assert!((sheet2.zoom - 2.0).abs() < f32::EPSILON);
    assert!((sheet2.view.zoom - 2.0).abs() < f32::EPSILON);

    // PANE frozen first row and column.
    assert_eq!(sheet2.frozen_rows, 1);
    assert_eq!(sheet2.frozen_cols, 1);
    assert_eq!(sheet2.view.pane.frozen_rows, 1);
    assert_eq!(sheet2.view.pane.frozen_cols, 1);
    assert_eq!(sheet2.view.pane.top_left_cell, Some(CellRef::new(1, 1))); // B2
    assert_eq!(sheet2.view.pane.x_split, None);
    assert_eq!(sheet2.view.pane.y_split, None);

    // WINDOW2 flags in the fixture clear showGridLines/showHeadings/showZeros.
    assert!(!sheet2.view.show_grid_lines);
    assert!(!sheet2.view.show_headings);
    assert!(!sheet2.view.show_zeros);

    // SELECTION active cell C3.
    let selection = sheet2.selection().expect("selection missing");
    assert_eq!(selection.active_cell, CellRef::new(2, 2)); // C3
    assert_eq!(selection.ranges, vec![Range::new(CellRef::new(2, 2), CellRef::new(2, 2))]);
}
