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
