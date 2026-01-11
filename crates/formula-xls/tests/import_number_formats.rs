use std::io::Write;

use formula_model::{CellRef, CellValue, DateSystem, Range};

mod xls_fixture_builder;

fn import_fixture(bytes: &[u8]) -> formula_xls::XlsImportResult {
    let mut tmp = tempfile::NamedTempFile::new().expect("temp file");
    tmp.write_all(bytes).expect("write xls bytes");
    formula_xls::import_xls_path(tmp.path()).expect("import xls")
}

#[test]
fn imports_biff_number_formats_and_formatted_blanks() {
    let bytes = xls_fixture_builder::build_number_format_fixture_xls(false);
    let result = import_fixture(&bytes);

    assert_eq!(result.workbook.date_system, DateSystem::Excel1900);

    let sheet = result
        .workbook
        .sheet_by_name("Formats")
        .expect("Formats missing");

    let cells = [
        ("A1", "$#,##0.00"),
        ("A2", "0.00%"),
        ("A3", "m/d/yy"),
        ("A4", "h:mm:ss"),
        ("A5", "[h]:mm:ss"),
    ];

    for (a1, expected_fmt) in cells {
        let cell_ref = CellRef::from_a1(a1).unwrap();
        let cell = sheet.cell(cell_ref).unwrap_or_else(|| panic!("{a1} missing"));
        assert!(
            !matches!(cell.value, CellValue::Empty),
            "{a1} unexpectedly empty"
        );

        let fmt = result
            .workbook
            .styles
            .get(cell.style_id)
            .and_then(|s| s.number_format.as_deref());
        assert_eq!(fmt, Some(expected_fmt), "number format mismatch for {a1}");
    }

    // A6 exists as a BLANK record with a non-General XF applied.
    let a6_ref = CellRef::from_a1("A6").unwrap();
    let a6 = sheet.cell(a6_ref).expect("A6 missing (formatted blank)");
    assert!(matches!(a6.value, CellValue::Empty));
    assert_ne!(a6.style_id, 0);
    let a6_fmt = result
        .workbook
        .styles
        .get(a6.style_id)
        .and_then(|s| s.number_format.as_deref());
    assert_eq!(a6_fmt, Some("0.00%"));
}

#[test]
fn respects_1904_date_system_flag() {
    let bytes = xls_fixture_builder::build_number_format_fixture_xls(true);
    let result = import_fixture(&bytes);
    assert_eq!(result.workbook.date_system, DateSystem::Excel1904);
}

#[test]
fn applies_formatted_blank_styles_to_merged_cell_anchors() {
    let bytes = xls_fixture_builder::build_merged_formatted_blank_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("MergedFmt")
        .expect("MergedFmt missing");

    let merge_range = Range::from_a1("A1:B1").unwrap();
    assert!(
        sheet
            .merged_regions
            .iter()
            .any(|region| region.range == merge_range),
        "missing expected merged range A1:B1"
    );

    // The BIFF BLANK record is for B1, but the model stores formatting on the
    // merged-region anchor (A1).
    let a1 = CellRef::from_a1("A1").unwrap();
    let b1 = CellRef::from_a1("B1").unwrap();

    let a1_cell = sheet.cell(a1).expect("A1 missing (formatted blank anchor)");
    let b1_cell = sheet.cell(b1).expect("B1 missing (merged)");
    assert!(matches!(a1_cell.value, CellValue::Empty));
    assert_eq!(a1_cell.style_id, b1_cell.style_id);
    assert_ne!(a1_cell.style_id, 0);

    let fmt = result
        .workbook
        .styles
        .get(a1_cell.style_id)
        .and_then(|s| s.number_format.as_deref());
    assert_eq!(fmt, Some("0.00%"));
}

#[test]
fn prefers_anchor_xf_when_merged_cells_have_conflicting_formats() {
    let bytes = xls_fixture_builder::build_merged_conflicting_blank_formats_fixture_xls();
    let result = import_fixture(&bytes);

    let sheet = result
        .workbook
        .sheet_by_name("MergedFmtConflict")
        .expect("MergedFmtConflict missing");

    // Formatting should follow the anchor cell (A1), not the non-anchor cell (B1).
    let a1 = CellRef::from_a1("A1").unwrap();
    let a1_cell = sheet.cell(a1).expect("A1 missing");
    assert!(matches!(a1_cell.value, CellValue::Empty));

    let fmt = result
        .workbook
        .styles
        .get(a1_cell.style_id)
        .and_then(|s| s.number_format.as_deref());
    assert_eq!(fmt, Some("0.00%"));
}

#[test]
fn date_system_affects_rendered_dates() {
    fn render_a3(result: &formula_xls::XlsImportResult) -> String {
        let sheet = result
            .workbook
            .sheet_by_name("Formats")
            .expect("Formats missing");
        let cell_ref = CellRef::from_a1("A3").unwrap();
        let cell = sheet.cell(cell_ref).expect("A3 missing");
        let serial = match &cell.value {
            CellValue::Number(v) => *v,
            other => panic!("A3 expected number, got {other:?}"),
        };
        let fmt = result
            .workbook
            .styles
            .get(cell.style_id)
            .and_then(|s| s.number_format.as_deref());
        let options = result
            .workbook
            .format_options(formula_format::Locale::en_us());
        formula_format::format_value(formula_format::Value::Number(serial), fmt, &options).text
    }

    let bytes_1900 = xls_fixture_builder::build_number_format_fixture_xls(false);
    let result_1900 = import_fixture(&bytes_1900);
    assert_eq!(render_a3(&result_1900), "7/16/23");

    let bytes_1904 = xls_fixture_builder::build_number_format_fixture_xls(true);
    let result_1904 = import_fixture(&bytes_1904);
    assert_eq!(render_a3(&result_1904), "7/17/27");
}
