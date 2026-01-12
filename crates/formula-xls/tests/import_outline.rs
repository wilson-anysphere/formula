use std::io::Write;

use formula_model::OutlineEntry;

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_row_and_col_outline_metadata() {
    let bytes = xls_fixture_builder::build_outline_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("Outline")
        .expect("Outline sheet missing");

    // WSBOOL-derived sheet-level outline settings.
    assert!(sheet.outline.pr.summary_below);
    assert!(sheet.outline.pr.summary_right);
    assert!(sheet.outline.pr.show_outline_symbols);

    // Levels (1-based indexing in the model):
    // - Rows 2-3 are detail rows at level 1.
    // - Row 4 is the collapsed summary row.
    assert_eq!(sheet.outline.rows.entry(2).level, 1);
    assert_eq!(sheet.outline.rows.entry(3).level, 1);
    assert_eq!(sheet.outline.rows.entry(4).level, 0);
    assert!(sheet.outline.rows.entry(4).collapsed);

    // Columns B-C (2-3) are detail columns at level 1; column D (4) is the collapsed summary col.
    assert_eq!(sheet.outline.cols.entry(2).level, 1);
    assert_eq!(sheet.outline.cols.entry(3).level, 1);
    assert_eq!(sheet.outline.cols.entry(4).level, 0);
    assert!(sheet.outline.cols.entry(4).collapsed);

    // Detail rows/cols are hidden by the collapsed outline group, not explicitly user-hidden.
    assert_eq!(sheet.outline.rows.entry(2).hidden.user, false);
    assert_eq!(sheet.outline.rows.entry(3).hidden.user, false);
    assert_eq!(sheet.outline.cols.entry(2).hidden.user, false);
    assert_eq!(sheet.outline.cols.entry(3).hidden.user, false);

    // Derive outline-hidden state from collapsed summary rows/cols.
    let mut outline = sheet.outline.clone();
    outline.recompute_outline_hidden_rows();
    outline.recompute_outline_hidden_cols();

    let hidden_row_2 = outline.rows.entry(2);
    let hidden_row_3 = outline.rows.entry(3);
    let summary_row_4 = outline.rows.entry(4);
    assert_eq!(
        hidden_row_2,
        OutlineEntry {
            level: 1,
            hidden: formula_model::HiddenState {
                user: false,
                outline: true,
                filter: false,
            },
            collapsed: false,
        }
    );
    assert!(hidden_row_3.hidden.outline);
    assert!(!summary_row_4.hidden.outline);

    assert!(outline.cols.entry(2).hidden.outline);
    assert!(outline.cols.entry(3).hidden.outline);
    assert!(!outline.cols.entry(4).hidden.outline);
}
