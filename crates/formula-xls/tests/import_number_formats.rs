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
