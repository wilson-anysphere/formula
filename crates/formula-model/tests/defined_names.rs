use formula_model::{
    validate_defined_name, DefinedNameError, DefinedNameScope, DefinedNameValidationError,
    Workbook, EXCEL_DEFINED_NAME_MAX_LEN, XLNM_FILTER_DATABASE, XLNM_PRINT_AREA, XLNM_PRINT_TITLES,
};

#[test]
fn validate_defined_name_accepts_common_and_builtin_names() {
    for name in [
        "MyRange",
        "_MyRange",
        "Name1",
        "Name.With.Dots",
        r"\BackslashStart",
        XLNM_PRINT_AREA,
        XLNM_PRINT_TITLES,
        XLNM_FILTER_DATABASE,
    ] {
        assert_eq!(
            validate_defined_name(name),
            Ok(()),
            "name should be valid: {name}"
        );
    }
}

#[test]
fn validate_defined_name_rejects_invalid_names() {
    assert_eq!(
        validate_defined_name(""),
        Err(DefinedNameValidationError::Empty)
    );
    assert_eq!(
        validate_defined_name("   "),
        Err(DefinedNameValidationError::Empty)
    );
    assert_eq!(
        validate_defined_name("A1"),
        Err(DefinedNameValidationError::LooksLikeCellReference)
    );
    assert_eq!(
        validate_defined_name("r1c1"),
        Err(DefinedNameValidationError::LooksLikeCellReference)
    );
    assert_eq!(
        validate_defined_name("1Name"),
        Err(DefinedNameValidationError::InvalidStartCharacter('1'))
    );
    assert_eq!(
        validate_defined_name(".Name"),
        Err(DefinedNameValidationError::InvalidStartCharacter('.'))
    );
    assert_eq!(
        validate_defined_name("My Name"),
        Err(DefinedNameValidationError::InvalidCharacter { ch: ' ', index: 2 })
    );

    let too_long = "a".repeat(EXCEL_DEFINED_NAME_MAX_LEN + 1);
    assert!(matches!(
        validate_defined_name(&too_long),
        Err(DefinedNameValidationError::TooLong { .. })
    ));
}

#[test]
fn defined_name_uniqueness_is_case_insensitive_and_scoped() {
    let mut wb = Workbook::new();
    let sheet1 = wb.add_sheet("Sheet1").unwrap();
    let sheet2 = wb.add_sheet("Sheet2").unwrap();

    let _workbook_id = wb
        .create_defined_name(
            DefinedNameScope::Workbook,
            "MyName",
            "Sheet1!A1",
            None,
            false,
            None,
        )
        .unwrap();

    assert_eq!(
        wb.create_defined_name(
            DefinedNameScope::Workbook,
            "myname",
            "Sheet1!B2",
            None,
            false,
            None,
        ),
        Err(DefinedNameError::DuplicateName)
    );

    // Same name is allowed in a different scope.
    wb.create_defined_name(
        DefinedNameScope::Sheet(sheet1),
        "MYNAME",
        "A1",
        None,
        false,
        None,
    )
    .unwrap();
    wb.create_defined_name(
        DefinedNameScope::Sheet(sheet2),
        "MyName",
        "A1",
        None,
        false,
        None,
    )
    .unwrap();
}

#[test]
fn defined_name_uniqueness_is_case_insensitive_for_unicode_text() {
    let mut wb = Workbook::new();
    wb.add_sheet("Sheet1").unwrap();

    wb.create_defined_name(
        DefinedNameScope::Workbook,
        "Straße",
        "Sheet1!A1",
        None,
        false,
        None,
    )
    .unwrap();

    // Uses Unicode-aware uppercasing: ß -> SS.
    assert_eq!(
        wb.create_defined_name(
            DefinedNameScope::Workbook,
            "STRASSE",
            "Sheet1!B2",
            None,
            false,
            None,
        ),
        Err(DefinedNameError::DuplicateName)
    );

    assert!(wb
        .get_defined_name(DefinedNameScope::Workbook, "STRASSE")
        .is_some());
}

#[test]
fn rename_sheet_rewrites_defined_name_refers_to() {
    let mut wb = Workbook::new();
    let sheet_id = wb.add_sheet("Sheet1").unwrap();

    wb.create_defined_name(
        DefinedNameScope::Workbook,
        "MyRange",
        "=Sheet1!$A$1:$A$3",
        None,
        false,
        None,
    )
    .unwrap();

    wb.create_defined_name(
        DefinedNameScope::Sheet(sheet_id),
        "LocalName",
        "Sheet1!B2",
        None,
        false,
        None,
    )
    .unwrap();

    wb.rename_sheet(sheet_id, "My Sheet").unwrap();

    let global = wb
        .get_defined_name(DefinedNameScope::Workbook, "myrange")
        .unwrap();
    assert_eq!(global.refers_to, "'My Sheet'!$A$1:$A$3");

    let local = wb
        .get_defined_name(DefinedNameScope::Sheet(sheet_id), "LocalName")
        .unwrap();
    assert_eq!(local.refers_to, "'My Sheet'!B2");
}
