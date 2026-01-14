use formula_model::{sanitize_sheet_name, validate_sheet_name, EXCEL_MAX_SHEET_NAME_LEN};

#[test]
fn removes_excel_forbidden_characters() {
    let sanitized = sanitize_sheet_name("  'Bad:Name\\With/Invalid?*Chars[here]'  ");
    assert_eq!(sanitized, "BadNameWithInvalidCharshere");
    assert!(validate_sheet_name(&sanitized).is_ok());
}

#[test]
fn truncates_to_excel_max_len_in_utf16_code_units() {
    // 🙂 is a non-BMP character, so it counts as 2 UTF-16 code units in Excel.
    let input = format!("{}🙂b", "a".repeat(EXCEL_MAX_SHEET_NAME_LEN - 2));
    let sanitized = sanitize_sheet_name(&input);
    let expected = format!("{}🙂", "a".repeat(EXCEL_MAX_SHEET_NAME_LEN - 2));
    assert_eq!(sanitized, expected);
    assert_eq!(sanitized.encode_utf16().count(), EXCEL_MAX_SHEET_NAME_LEN);
    assert!(validate_sheet_name(&sanitized).is_ok());
}

#[test]
fn falls_back_to_sheet1_for_empty_or_whitespace_names() {
    assert_eq!(sanitize_sheet_name(""), "Sheet1");
    assert_eq!(sanitize_sheet_name("   "), "Sheet1");
    assert_eq!(sanitize_sheet_name("[]"), "Sheet1");
    assert!(validate_sheet_name(&sanitize_sheet_name("[]")).is_ok());
}

#[test]
fn always_produces_a_valid_sheet_name() {
    for input in [
        "",
        "   ",
        "[]",
        "  [Data]  ",
        "'Leading",
        "Trailing'",
        "this name is way too long for excel sheets and must be truncated",
        "a🙂b🙂c🙂d🙂e🙂f🙂g🙂h🙂i🙂j🙂k🙂l🙂m🙂n🙂o🙂p",
    ] {
        let sanitized = sanitize_sheet_name(input);
        assert!(
            validate_sheet_name(&sanitized).is_ok(),
            "input={input:?}, sanitized={sanitized:?}"
        );
    }
}
