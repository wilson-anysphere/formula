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
    assert_eq!(cell_format_code(Some("__builtin_numFmtId:999")), "G");
}

#[test]
fn cell_format_code_builtin_datetime_placeholder_is_datetime() {
    let code = cell_format_code(Some("__builtin_numFmtId:14"));
    assert!(
        code.starts_with('D') || code.starts_with('T'),
        "expected date/time classification for built-in 14, got {code}"
    );
}

#[test]
fn cell_format_code_reserved_datetime_builtin_placeholder_is_datetime() {
    let code = cell_format_code(Some("__builtin_numFmtId:50"));
    assert!(
        code.starts_with('D') || code.starts_with('T'),
        "expected date/time classification for reserved built-in 50, got {code}"
    );
}
