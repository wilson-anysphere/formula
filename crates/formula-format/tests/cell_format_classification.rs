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

#[test]
fn cell_format_code_classifies_system_long_date_tokens_as_date_not_currency() {
    // Excel uses special "system" tokens like `[$-F800]` for locale-dependent long date formats.
    // These are *not* currency tokens and should not force currency classification.
    let code = cell_format_code(Some("[$-F800]dddd, mmmm dd, yyyy"));
    assert!(
        code.starts_with('D'),
        "expected date classification (D*), got {code:?}"
    );
}

#[test]
fn cell_format_code_treats_day_first_dates_like_month_first_for_cell_classification() {
    let mdy = cell_format_code(Some("m/d/yyyy"));
    let dmy = cell_format_code(Some("dd/mm/yyyy"));
    let dmy_single = cell_format_code(Some("d/m/yyyy"));

    assert!(mdy.starts_with('D'), "expected date classification, got {mdy:?}");
    assert_eq!(dmy, mdy);
    assert_eq!(dmy_single, mdy);
}

#[test]
fn cell_format_code_treats_hh_mm_like_h_mm_for_time_classification() {
    let h = cell_format_code(Some("h:mm"));
    let hh = cell_format_code(Some("hh:mm"));

    assert!(h.starts_with('T'), "expected time classification, got {h:?}");
    assert_eq!(hh, h);
}
