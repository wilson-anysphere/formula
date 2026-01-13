use std::collections::BTreeSet;
use std::io::Write;

use formula_model::{ManualPageBreaks, Orientation, PageMargins, PageSetup, Scaling};

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_page_setup_for_sanitized_sheet_name() {
    let bytes = xls_fixture_builder::build_page_setup_sanitized_sheet_name_fixture_xls();
    let result = import_fixture(&bytes);
    let workbook = result.workbook;

    // `Bad:Name` is invalid per Excel rules; importer should sanitize `:` -> `_`.
    assert!(
        workbook.sheet_by_name("Bad_Name").is_some(),
        "expected sanitized sheet name `Bad_Name`; sheets={:?}",
        workbook.sheets.iter().map(|s| &s.name).collect::<Vec<_>>()
    );

    let settings = workbook.sheet_print_settings_by_name("Bad_Name");
    assert_eq!(
        settings.page_setup,
        PageSetup {
            orientation: Orientation::Landscape,
            paper_size: formula_model::PaperSize { code: 9 },
            margins: PageMargins {
                left: 1.11,
                right: 2.22,
                top: 3.33,
                bottom: 4.44,
                header: 0.55,
                footer: 0.66,
            },
            scaling: Scaling::Percent(123),
        }
    );

    let mut expected_row_breaks: BTreeSet<u32> = BTreeSet::new();
    expected_row_breaks.insert(1);
    expected_row_breaks.insert(4);
    let mut expected_col_breaks: BTreeSet<u32> = BTreeSet::new();
    expected_col_breaks.insert(2);
    assert_eq!(
        settings.manual_page_breaks,
        ManualPageBreaks {
            row_breaks_after: expected_row_breaks,
            col_breaks_after: expected_col_breaks,
        }
    );

    // Ensure the print settings were not stored under the *original* (invalid) BIFF name.
    let original_name_settings = workbook.sheet_print_settings_by_name("Bad:Name");
    assert_eq!(original_name_settings, formula_model::SheetPrintSettings::new("Bad:Name"));
}
