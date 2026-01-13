use std::collections::BTreeSet;
use std::io::Write;

use formula_model::{
    ManualPageBreaks, Orientation, PageMargins, PageSetup, PaperSize, Scaling, SheetPrintSettings,
};

mod common;

use common::xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

fn manual_breaks(row_break_after: u32, col_break_after: u32) -> ManualPageBreaks {
    let mut row_breaks_after = BTreeSet::new();
    row_breaks_after.insert(row_break_after);
    let mut col_breaks_after = BTreeSet::new();
    col_breaks_after.insert(col_break_after);
    ManualPageBreaks {
        row_breaks_after,
        col_breaks_after,
    }
}

#[test]
fn imports_page_setup_margins_and_breaks_per_sheet_from_biff_multisheet() {
    let bytes = xls_fixture_builder::build_page_setup_multisheet_fixture_xls();
    let result = import_fixture(&bytes);
    let workbook = &result.workbook;

    assert_eq!(
        workbook.print_settings.sheets.len(),
        2,
        "expected print settings for both sheets; got {:?}; warnings={:?}",
        workbook.print_settings.sheets,
        result.warnings
    );
    assert_eq!(workbook.print_settings.sheets[0].sheet_name, "First");
    assert_eq!(workbook.print_settings.sheets[1].sheet_name, "Second");

    let mut expected_first = SheetPrintSettings::new("First");
    expected_first.page_setup = PageSetup {
        orientation: Orientation::Landscape,
        paper_size: PaperSize { code: 9 },
        margins: PageMargins {
            left: 0.5,
            right: 1.0,
            top: 1.5,
            bottom: 2.0,
            header: 0.125,
            footer: 0.875,
        },
        scaling: Scaling::Percent(80),
    };
    expected_first.manual_page_breaks = manual_breaks(4, 1);
    assert_eq!(
        workbook.sheet_print_settings_by_name("First"),
        expected_first,
        "First sheet settings mismatch; warnings={:?}",
        result.warnings
    );

    let mut expected_second = SheetPrintSettings::new("Second");
    expected_second.page_setup = PageSetup {
        orientation: Orientation::Portrait,
        paper_size: PaperSize { code: 1 },
        margins: PageMargins {
            left: 0.375,
            right: 0.625,
            top: 1.125,
            bottom: 1.875,
            header: 0.25,
            footer: 0.75,
        },
        scaling: Scaling::Percent(120),
    };
    expected_second.manual_page_breaks = manual_breaks(9, 3);
    assert_eq!(
        workbook.sheet_print_settings_by_name("Second"),
        expected_second,
        "Second sheet settings mismatch; warnings={:?}",
        result.warnings
    );
}

