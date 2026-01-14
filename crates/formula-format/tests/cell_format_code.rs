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
