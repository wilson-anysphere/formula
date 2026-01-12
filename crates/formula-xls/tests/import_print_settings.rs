use std::io::Write;

use formula_model::{ColRange, Range, RowRange};

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_print_settings_from_biff_builtin_defined_names() {
    let bytes = xls_fixture_builder::build_defined_names_builtins_fixture_xls();
    let result = import_fixture(&bytes);
    let workbook = result.workbook;

    let sheet1_settings = workbook.sheet_print_settings_by_name("Sheet1");
    assert_eq!(
        sheet1_settings.print_area,
        Some(vec![
            Range::from_a1("A1:A2").unwrap(),
            Range::from_a1("C1:C2").unwrap()
        ])
    );

    let sheet2_settings = workbook.sheet_print_settings_by_name("Sheet2");
    let titles = sheet2_settings
        .print_titles
        .expect("expected print_titles for Sheet2");
    assert_eq!(titles.repeat_rows, Some(RowRange { start: 0, end: 0 }));
    assert_eq!(titles.repeat_cols, Some(ColRange { start: 0, end: 0 }));
}

