use base64::{engine::general_purpose::STANDARD, Engine as _};
use formula_xlsx::print::{
    read_workbook_print_settings, write_workbook_print_settings, CellRange, ColRange, Orientation,
    PrintTitles, RowRange, Scaling,
};

fn load_fixture_xlsx() -> Vec<u8> {
    let fixture_path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/print-settings.xlsx.base64");
    let data = std::fs::read_to_string(&fixture_path).expect("fixture base64 should be readable");
    let cleaned: String = data.lines().map(str::trim).collect();
    STANDARD
        .decode(cleaned.as_bytes())
        .expect("fixture base64 should decode")
}

#[test]
fn preserves_print_settings_and_allows_updates() {
    let original = load_fixture_xlsx();
    let settings = read_workbook_print_settings(&original).expect("read print settings");

    assert_eq!(settings.sheets.len(), 1);
    let sheet = &settings.sheets[0];
    assert_eq!(sheet.sheet_name, "Sheet1");

    assert_eq!(
        sheet.print_area.as_deref(),
        Some(
            &[CellRange {
                start_row: 1,
                end_row: 10,
                start_col: 1,
                end_col: 4
            }][..]
        )
    );

    assert_eq!(
        sheet.print_titles,
        Some(PrintTitles {
            repeat_rows: Some(RowRange { start: 1, end: 1 }),
            repeat_cols: Some(ColRange { start: 1, end: 2 }),
        })
    );

    assert_eq!(sheet.page_setup.orientation, Orientation::Landscape);
    assert_eq!(sheet.page_setup.paper_size.code, 9);
    assert_eq!(
        sheet.page_setup.scaling,
        Scaling::FitTo {
            width: 1,
            height: 0
        }
    );

    assert!(sheet.manual_page_breaks.row_breaks_after.contains(&5));
    assert!(sheet.manual_page_breaks.col_breaks_after.contains(&2));

    // Update a couple of fields and ensure they survive a write/read.
    let mut updated_settings = settings.clone();
    let updated_sheet = &mut updated_settings.sheets[0];
    updated_sheet.print_area = Some(vec![CellRange {
        start_row: 2,
        end_row: 5,
        start_col: 2,
        end_col: 3,
    }]);
    updated_sheet.page_setup.scaling = Scaling::Percent(80);
    updated_sheet.manual_page_breaks.row_breaks_after.clear();
    updated_sheet.manual_page_breaks.row_breaks_after.insert(3);

    let rewritten =
        write_workbook_print_settings(&original, &updated_settings).expect("write print settings");
    let reread = read_workbook_print_settings(&rewritten).expect("re-read print settings");

    assert_eq!(reread.sheets.len(), 1);
    let sheet = &reread.sheets[0];

    assert_eq!(
        sheet.print_area.as_deref(),
        Some(
            &[CellRange {
                start_row: 2,
                end_row: 5,
                start_col: 2,
                end_col: 3
            }][..]
        )
    );

    assert_eq!(sheet.page_setup.scaling, Scaling::Percent(80));
    assert!(sheet.manual_page_breaks.row_breaks_after.contains(&3));
    assert!(!sheet.manual_page_breaks.row_breaks_after.contains(&5));
}
