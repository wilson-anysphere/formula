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
        code.starts_with('D'),
        "expected date/time classification for built-in 14, got {code}"
    );
}

#[test]
fn cell_format_code_reserved_datetime_builtin_placeholder_is_datetime() {
    let code = cell_format_code(Some("__builtin_numFmtId:50"));
    assert!(
        code.starts_with('D'),
        "expected date/time classification for reserved built-in 50, got {code}"
    );
}

#[test]
fn cell_format_code_classifies_thousands_separated_numbers_as_n() {
    assert_eq!(cell_format_code(Some("#,##0")), "N0");
    assert_eq!(cell_format_code(Some("#,##0.00")), "N2");

    // Placeholder variants for Excel built-ins 3 and 4.
    assert_eq!(cell_format_code(Some("__builtin_numFmtId:3")), "N0");
    assert_eq!(cell_format_code(Some("__builtin_numFmtId:4")), "N2");

    // No grouping => fixed classification.
    assert_eq!(cell_format_code(Some("0.00")), "F2");
}

#[test]
fn cell_format_code_uses_positive_section_for_grouping_detection() {
    // Grouping commas in non-positive sections should not affect the CELL("format") classification.
    assert_eq!(cell_format_code(Some("0;#,##0")), "F0");
    assert_eq!(cell_format_code(Some("0;0;#,##0")), "F0");

    // Conditional sections: Excel selects the first matching condition, then the first
    // unconditional section as an "else". CELL("format") uses the selected *positive* section.
    assert_eq!(cell_format_code(Some("[<0]#,##0;0")), "F0");
    assert_eq!(cell_format_code(Some("[>=0]#,##0;0")), "N0");
}

#[test]
fn cell_format_code_ignores_commas_in_literals_escapes_and_brackets() {
    // Commas inside quoted literals are not thousands separators.
    assert_eq!(cell_format_code(Some(r#"0","0"#)), "F0");

    // Escaped commas are rendered literally and should not count as grouping.
    assert_eq!(cell_format_code(Some(r#"0\,00"#)), "F0");

    // Commas inside bracket tokens should be ignored for grouping detection.
    assert_eq!(cell_format_code(Some("[foo,bar]0.00")), "F2");
}

#[test]
fn cell_format_code_ignores_commas_in_layout_tokens() {
    // `_X` / `*X` layout tokens consume their operand, which is not rendered literally. A comma used
    // as a layout operand is not a thousands separator.
    assert_eq!(cell_format_code(Some("0_,0")), "F0");
    assert_eq!(cell_format_code(Some("0*,0")), "F0");
}
