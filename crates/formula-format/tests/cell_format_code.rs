use formula_format::cell_format_code;

#[test]
fn cell_format_code_detects_multi_char_currency_bracket_tokens() {
    assert_eq!(cell_format_code(Some("[$USD-409]#,##0.00")), "C2");
    assert_eq!(cell_format_code(Some("[$CAD-1009]0")), "C0");

    // `[$-409]` is a locale override without an explicit currency symbol/code.
    assert!(
        !cell_format_code(Some("[$-409]0")).starts_with('C'),
        "locale-only bracket tokens should not be treated as currency"
    );
}

#[test]
fn cell_format_code_unknown_builtin_placeholder_falls_back_to_general() {
    // Unknown placeholders should behave like General, and must not be misclassified as date/time
    // based on the placeholder text itself.
    assert_eq!(cell_format_code(Some("__builtin_numFmtId:999")), "G");
}

#[test]
fn cell_format_code_reserved_datetime_builtin_placeholder_is_datetime() {
    // Excel reserves additional built-in ids for locale-specific date/time formats (e.g. 50-58).
    // Even though they are not in the 0-49 built-in table, they should still classify as
    // date/time rather than falling back to number formats.
    assert!(
        cell_format_code(Some("__builtin_numFmtId:50")).starts_with('D'),
        "expected date/time classification for reserved date/time placeholder id"
    );
}
