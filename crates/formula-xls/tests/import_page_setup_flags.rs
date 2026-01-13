use std::io::Write;

use formula_model::{Orientation, PaperSize, Scaling};

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_setup_f_no_pls_ignores_paper_size_orientation_and_scale_but_keeps_header_footer_margins()
{
    let bytes = xls_fixture_builder::build_page_setup_flags_nopls_fixture_xls();
    let result = import_fixture(&bytes);
    let workbook = result.workbook;

    let settings = workbook.sheet_print_settings_by_name("PageSetupNoPls");
    assert_eq!(settings.page_setup.orientation, Orientation::Portrait);
    assert_eq!(settings.page_setup.paper_size, PaperSize::LETTER);
    assert_eq!(settings.page_setup.scaling, Scaling::Percent(100));

    // Header/footer margins are stored in SETUP.numHdr/numFtr; these should still import even when
    // `fNoPls=1` causes other fields to be ignored.
    assert_eq!(settings.page_setup.margins.header, 0.5);
    assert_eq!(settings.page_setup.margins.footer, 0.6);
}

#[test]
fn imports_setup_f_no_orient_ignores_f_portrait() {
    let bytes = xls_fixture_builder::build_page_setup_flags_noorient_fixture_xls();
    let result = import_fixture(&bytes);
    let workbook = result.workbook;

    let settings = workbook.sheet_print_settings_by_name("PageSetupNoOrient");
    assert_eq!(settings.page_setup.orientation, Orientation::Portrait);

    // Other SETUP fields are still respected when `fNoPls=0`.
    assert_eq!(settings.page_setup.paper_size, PaperSize::A4);
    assert_eq!(settings.page_setup.scaling, Scaling::Percent(80));
}
