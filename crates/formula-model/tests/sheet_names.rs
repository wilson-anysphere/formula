use formula_model::{RenameSheetError, SheetNameError, Workbook};

#[test]
fn valid_sheet_names_are_accepted() {
    let mut workbook = Workbook::new();
    workbook.add_sheet("Sheet1").unwrap();
    workbook.add_sheet("My Sheet").unwrap();
    workbook.add_sheet("Résumé").unwrap();
    workbook.add_sheet("こんにちは").unwrap();
    workbook.add_sheet("O'Brien").unwrap();
}

#[test]
fn rejects_each_excel_forbidden_character() {
    for ch in [':', '\\', '/', '?', '*', '[', ']'] {
        let mut workbook = Workbook::new();
        let name = format!("Bad{ch}Name");
        assert_eq!(
            workbook.add_sheet(name),
            Err(SheetNameError::InvalidCharacter(ch))
        );
    }
}

#[test]
fn length_boundaries_match_excel() {
    let mut workbook = Workbook::new();

    workbook.add_sheet("A").unwrap();

    let max = "a".repeat(formula_model::EXCEL_MAX_SHEET_NAME_LEN);
    workbook.add_sheet(max).unwrap();

    let too_long = "a".repeat(formula_model::EXCEL_MAX_SHEET_NAME_LEN + 1);
    assert_eq!(workbook.add_sheet(too_long), Err(SheetNameError::TooLong));
}

#[test]
fn length_limit_counts_utf16_code_units() {
    // 🙂 is a non-BMP character, so it counts as 2 UTF-16 code units in Excel.
    let mut workbook = Workbook::new();
    let name = format!(
        "{}🙂",
        "a".repeat(formula_model::EXCEL_MAX_SHEET_NAME_LEN - 1)
    );
    assert_eq!(workbook.add_sheet(name), Err(SheetNameError::TooLong));
}

#[test]
fn rejects_blank_or_whitespace_only_names() {
    let mut workbook = Workbook::new();
    assert_eq!(workbook.add_sheet(""), Err(SheetNameError::EmptyName));
    assert_eq!(workbook.add_sheet("   "), Err(SheetNameError::EmptyName));
}

#[test]
fn rejects_leading_or_trailing_apostrophe() {
    let mut workbook = Workbook::new();
    assert_eq!(
        workbook.add_sheet("'Leading"),
        Err(SheetNameError::LeadingOrTrailingApostrophe)
    );
    assert_eq!(
        workbook.add_sheet("Trailing'"),
        Err(SheetNameError::LeadingOrTrailingApostrophe)
    );
}

#[test]
fn detects_duplicates_case_insensitively_on_add_and_rename() {
    let mut workbook = Workbook::new();
    let data = workbook.add_sheet("Data").unwrap();
    let summary = workbook.add_sheet("Summary").unwrap();

    assert_eq!(
        workbook.add_sheet("data"),
        Err(SheetNameError::DuplicateName)
    );

    assert_eq!(
        workbook.rename_sheet(summary, "DATA"),
        Err(RenameSheetError::InvalidName(SheetNameError::DuplicateName))
    );

    // Renaming the same sheet to a different case is allowed.
    workbook.rename_sheet(data, "data").unwrap();
    assert_eq!(workbook.sheets[0].name, "data");
}

#[test]
fn detects_duplicates_with_unicode_case_insensitive_matching() {
    let mut workbook = Workbook::new();
    workbook.add_sheet("Äbc").unwrap();
    assert_eq!(
        workbook.add_sheet("äbc"),
        Err(SheetNameError::DuplicateName)
    );
}
