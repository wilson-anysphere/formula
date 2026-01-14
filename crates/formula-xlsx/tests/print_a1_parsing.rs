use formula_xlsx::print::{
    parse_print_area_defined_name, parse_print_titles_defined_name, ColRange, PrintTitles, RowRange,
};

#[test]
fn print_defined_name_parsing_accepts_sheet_names_case_insensitively() {
    let area = parse_print_area_defined_name("Sheet1", "sheet1!$A$1:$B$2").unwrap();
    assert_eq!(area.len(), 1);

    let titles = parse_print_titles_defined_name("Sheet1", "sHeEt1!$1:$3,sheet1!$A:$B").unwrap();
    assert_eq!(
        titles,
        PrintTitles {
            repeat_rows: Some(RowRange { start: 1, end: 3 }),
            repeat_cols: Some(ColRange { start: 1, end: 2 }),
        }
    );
}

#[test]
fn print_defined_name_parsing_uses_unicode_case_insensitive_sheet_matching() {
    // German sharp s: Unicode uppercasing expands `ß` -> `SS`.
    let area = parse_print_area_defined_name("ß", "SS!$A$1:$B$2").unwrap();
    assert_eq!(area.len(), 1);
}

#[test]
fn print_defined_name_parsing_rejects_overflowing_column_letters() {
    // Extremely large column references should be rejected without panicking on integer overflow.
    assert!(
        parse_print_area_defined_name("Sheet1", "Sheet1!$ZZZZZZZ$1:$A$1").is_err(),
        "expected parse to fail for column letters that overflow u32"
    );
}

#[test]
fn print_titles_defined_name_parsing_accepts_row_and_col_cell_ranges() {
    // Some producers represent print titles using explicit cell ranges rather than row/col-only
    // references. Accept these best-effort.
    let titles = parse_print_titles_defined_name("Sheet1", "Sheet1!$A$1:$IV$1,Sheet1!$A$1:$A$10")
        .unwrap();
    assert_eq!(
        titles,
        PrintTitles {
            repeat_rows: Some(RowRange { start: 1, end: 1 }),
            repeat_cols: Some(ColRange { start: 1, end: 1 }),
        }
    );

    // But reject cell ranges that span both multiple rows and multiple columns (ambiguous).
    assert!(
        parse_print_titles_defined_name("Sheet1", "Sheet1!$A$1:$B$2").is_err(),
        "expected multi-row multi-col cell range to be rejected"
    );
}
