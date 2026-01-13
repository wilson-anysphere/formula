use std::io::Write;

use formula_model::{PageMargins, PaperSize, Scaling};

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_margins_without_setup_record() {
    let bytes = xls_fixture_builder::build_margins_without_setup_fixture_xls();
    let result = import_fixture(&bytes);

    let settings = result.workbook.sheet_print_settings_by_name("Sheet1");
    let page_setup = settings.page_setup;

    // Worksheet has no SETUP record, so all non-margin page setup options should remain default.
    assert_eq!(page_setup.orientation, Default::default());
    assert_eq!(page_setup.paper_size, PaperSize::default());
    assert_eq!(page_setup.scaling, Scaling::Percent(100));

    // Margins should come from the standalone margin records.
    assert_eq!(page_setup.margins.left, 1.25);
    assert_eq!(page_setup.margins.right, 1.5);
    assert_eq!(page_setup.margins.top, 0.5);
    assert_eq!(page_setup.margins.bottom, 2.25);

    // Header/footer margins live in the SETUP record; without SETUP they should remain default.
    let defaults = PageMargins::default();
    assert_eq!(page_setup.margins.header, defaults.header);
    assert_eq!(page_setup.margins.footer, defaults.footer);
}

