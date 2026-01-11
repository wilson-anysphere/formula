use std::io::Write;

use formula_model::{CellRef, CellValue, DateSystem};

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

