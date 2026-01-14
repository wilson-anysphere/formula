use formula_format::cell_format_code;

#[test]
fn cell_format_code_detects_currency_symbols_and_bracket_tokens() {
    // OOXML currency + locale token (common in XLSX/XLSB styles).
    let code = cell_format_code(Some("[$€-407]#,##0.00"));
    assert!(
        code.starts_with("C2"),
        "expected currency classification (C2*), got {code:?}"
    );

    // Currency symbol as a literal outside quotes.
    let code = cell_format_code(Some("€#,##0.00"));
    assert!(
        code.starts_with("C2"),
        "expected currency classification (C2*), got {code:?}"
    );

    // Locale override with no symbol should *not* be treated as currency.
    let code = cell_format_code(Some("[$-409]0.00"));
    assert_eq!(code, "F2");
}

