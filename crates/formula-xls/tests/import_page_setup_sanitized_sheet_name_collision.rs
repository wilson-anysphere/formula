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
fn imports_page_setup_when_sanitization_collides_with_another_sheet_name() {
    let bytes = xls_fixture_builder::build_page_setup_sanitized_sheet_name_collision_fixture_xls();
    let result = import_fixture(&bytes);
    let workbook = result.workbook;

    // Sheet 0: `Bad:Name` is invalid and sanitizes to `Bad_Name`.
    assert!(workbook.sheet_by_name("Bad_Name").is_some());
    // Sheet 1: original `Bad_Name` collides and is deduped.
    assert!(workbook.sheet_by_name("Bad_Name (2)").is_some());

    // The invalid BIFF name should never exist as a worksheet name.
    assert!(workbook.sheet_by_name("Bad:Name").is_none());

    assert_eq!(
        workbook
            .print_settings
            .sheets
            .iter()
            .map(|s| s.sheet_name.as_str())
            .collect::<Vec<_>>(),
        vec!["Bad_Name", "Bad_Name (2)"],
        "expected print settings entries to be keyed by the final workbook sheet names"
    );

    // Sheet 0 settings (matches `build_page_setup_sanitized_sheet_name_fixture_xls`).
    let settings0 = workbook.sheet_print_settings_by_name("Bad_Name");
    assert_eq!(
        settings0.page_setup,
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
    assert_eq!(
        settings0.manual_page_breaks,
        ManualPageBreaks {
            row_breaks_after: BTreeSet::from([1u32, 4u32]),
            col_breaks_after: BTreeSet::from([2u32]),
        }
    );

    // Sheet 1 settings (distinct values).
    let settings1 = workbook.sheet_print_settings_by_name("Bad_Name (2)");
    assert_eq!(
        settings1.page_setup,
        PageSetup {
            orientation: Orientation::Portrait,
            paper_size: formula_model::PaperSize { code: 1 },
            margins: PageMargins {
                left: 5.55,
                right: 6.66,
                top: 7.77,
                bottom: 8.88,
                header: 0.11,
                footer: 0.22,
            },
            scaling: Scaling::Percent(77),
        }
    );
    assert_eq!(
        settings1.manual_page_breaks,
        ManualPageBreaks {
            row_breaks_after: BTreeSet::from([2u32]),
            col_breaks_after: BTreeSet::from([1u32]),
        }
    );

    // Ensure print settings were not stored under the original invalid BIFF name.
    let original_name_settings = workbook.sheet_print_settings_by_name("Bad:Name");
    assert_eq!(original_name_settings, formula_model::SheetPrintSettings::new("Bad:Name"));
}

