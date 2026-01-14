use std::io::Cursor;

use formula_model::{Range, SheetPrintSettings, Workbook};
use formula_xlsx::print::{read_workbook_print_settings, write_workbook_print_settings, CellRange};

#[test]
fn writer_matches_print_settings_sheet_names_nfkc_case_insensitively(
) -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    workbook.add_sheet("Kelvin")?;

    // U+212A Kelvin sign should compare equal to ASCII 'K' under Excel-like semantics
    // (Unicode NFKC + case-insensitive).
    let mut sheet_settings = SheetPrintSettings::new("Kelvin");
    sheet_settings.print_area = Some(vec![Range::from_a1("A1")?]);
    workbook.print_settings.sheets = vec![sheet_settings];

    let mut buf = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&workbook, &mut buf)?;
    let bytes = buf.into_inner();

    let read = read_workbook_print_settings(&bytes)?;
    assert_eq!(read.sheets.len(), 1);
    let sheet = &read.sheets[0];
    assert_eq!(sheet.sheet_name, "Kelvin");
    assert_eq!(
        sheet.print_area.as_deref(),
        Some(
            &[CellRange {
                start_row: 1,
                end_row: 1,
                start_col: 1,
                end_col: 1,
            }][..]
        )
    );

    Ok(())
}

#[test]
fn print_writer_matches_sheet_names_nfkc_case_insensitively(
) -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    workbook.add_sheet("Kelvin")?;

    let mut buf = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&workbook, &mut buf)?;
    let original = buf.into_inner();

    let mut settings = read_workbook_print_settings(&original)?;
    assert_eq!(settings.sheets.len(), 1);

    settings.sheets[0].sheet_name = "Kelvin".to_string();
    settings.sheets[0].print_area = Some(vec![CellRange {
        start_row: 2,
        end_row: 3,
        start_col: 2,
        end_col: 3,
    }]);

    let rewritten = write_workbook_print_settings(&original, &settings)?;
    let reread = read_workbook_print_settings(&rewritten)?;

    assert_eq!(reread.sheets.len(), 1);
    let sheet = &reread.sheets[0];
    assert_eq!(sheet.sheet_name, "Kelvin");
    assert_eq!(
        sheet.print_area.as_deref(),
        Some(
            &[CellRange {
                start_row: 2,
                end_row: 3,
                start_col: 2,
                end_col: 3,
            }][..]
        )
    );

    Ok(())
}

#[test]
fn xlsx_document_print_settings_patches_match_sheet_names_nfkc_case_insensitively(
) -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    workbook.add_sheet("Kelvin")?;

    let mut buf = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&workbook, &mut buf)?;
    let original = buf.into_inner();

    let mut doc = formula_xlsx::load_from_bytes(&original)?;
    let mut sheet_settings = SheetPrintSettings::new("Kelvin");
    sheet_settings.print_area = Some(vec![Range::from_a1("A1")?]);
    doc.workbook.print_settings.sheets = vec![sheet_settings];

    let saved = doc.save_to_vec()?;
    let reread = read_workbook_print_settings(&saved)?;

    assert_eq!(reread.sheets.len(), 1);
    let sheet = &reread.sheets[0];
    assert_eq!(sheet.sheet_name, "Kelvin");
    assert_eq!(
        sheet.print_area.as_deref(),
        Some(
            &[CellRange {
                start_row: 1,
                end_row: 1,
                start_col: 1,
                end_col: 1,
            }][..]
        )
    );

    Ok(())
}

