use formula_format::cell_format_code;
use formula_format::{builtin_format_code, classify_cell_format, CellFormatClassification};

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
fn cell_format_code_recognizes_year_first_iso_dates() {
    let mdy = cell_format_code(Some("m/d/yyyy"));
    let iso_dash = cell_format_code(Some("yyyy-mm-dd"));
    let iso_slash = cell_format_code(Some("yyyy/m/d"));

    assert!(mdy.starts_with('D'), "expected date classification, got {mdy:?}");
    assert_eq!(iso_dash, mdy);
    assert_eq!(iso_slash, mdy);
}

#[test]
fn cell_format_code_treats_hh_mm_like_h_mm_for_time_classification() {
    let h = cell_format_code(Some("h:mm"));
    let hh = cell_format_code(Some("hh:mm"));

    assert_eq!(h, "D9", "expected Excel `h:mm` classification, got {h:?}");
    assert_eq!(hh, h);
}

#[test]
fn cell_format_code_treats_hh_mm_ss_like_h_mm_ss_for_time_classification() {
    let h = cell_format_code(Some("h:mm:ss"));
    let hh = cell_format_code(Some("hh:mm:ss"));

    assert_eq!(h, "D8", "expected Excel `h:mm:ss` classification, got {h:?}");
    assert_eq!(hh, h);
}

#[test]
fn cell_format_code_ignores_locale_override_tokens_for_datetime_classification() {
    let base_date = cell_format_code(Some("dd/mm/yyyy"));
    let with_locale = cell_format_code(Some("[$-409]dd/mm/yyyy"));

    assert!(base_date.starts_with('D'), "expected date classification, got {base_date:?}");
    assert_eq!(with_locale, base_date);

    let base_time = cell_format_code(Some("hh:mm"));
    let with_locale = cell_format_code(Some("[$-409]hh:mm"));
    assert_eq!(base_time, "D9", "expected Excel `h:mm` classification, got {base_time:?}");
    assert_eq!(with_locale, base_time);
}

#[test]
fn builtin_numeric_formats_0_11_map_to_expected_cell_format_codes() {
    let cases: &[(u16, &str)] = &[
        (0, "G"),
        (1, "F0"),
        (2, "F2"),
        (3, "F0"),
        (4, "F2"),
        (5, "C0"),
        (6, "C0"),
        (7, "C2"),
        (8, "C2"),
        (9, "P0"),
        (10, "P2"),
        (11, "S2"),
    ];

    for &(id, expected_code) in cases {
        let fmt = builtin_format_code(id).unwrap();
        let info = classify_cell_format(Some(fmt));
        assert_eq!(info.cell_format_code, expected_code, "id {id} ({fmt})");
    }

    // Currency formats: red + parentheses.
    let fmt6 = builtin_format_code(6).unwrap();
    let info6 = classify_cell_format(Some(fmt6));
    assert!(info6.negative_in_color, "expected [Red] in negative section");
    assert!(info6.negative_in_parentheses, "expected parentheses in negative section");

    let fmt8 = builtin_format_code(8).unwrap();
    let info8 = classify_cell_format(Some(fmt8));
    assert!(info8.negative_in_color, "expected [Red] in negative section");
    assert!(info8.negative_in_parentheses, "expected parentheses in negative section");
}

#[test]
fn accounting_formats_detect_parentheses_and_negative_color() {
    // Built-in accounting-style negatives (no currency symbol): 23–26.
    let cases: &[(u16, &str, bool)] = &[
        (23, "F0", false),
        (24, "F0", true),
        (25, "F2", false),
        (26, "F2", true),
    ];

    for &(id, expected_code, expects_color) in cases {
        let fmt = builtin_format_code(id).unwrap();
        let info = classify_cell_format(Some(fmt));
        assert_eq!(info.cell_format_code, expected_code, "id {id} ({fmt})");
        assert!(
            info.negative_in_parentheses,
            "expected parentheses in negative section for id {id} ({fmt})"
        );
        assert_eq!(
            info.negative_in_color, expects_color,
            "id {id} ({fmt})"
        );
    }

    // Accounting formats 41–44 (alignment underscores/fill).
    for id in 41u16..=44u16 {
        let fmt = builtin_format_code(id).unwrap();
        let info = classify_cell_format(Some(fmt));
        assert!(
            info.negative_in_parentheses,
            "expected parentheses in negative section for id {id} ({fmt})"
        );
        assert!(
            !info.negative_in_color,
            "did not expect a color token for id {id} ({fmt})"
        );
    }
}

#[test]
fn builtin_placeholder_inputs_match_builtin_strings() {
    for id in [0u16, 6u16, 9u16, 14u16, 41u16, 49u16] {
        let fmt = builtin_format_code(id).unwrap();
        let from_string = classify_cell_format(Some(fmt));
        let from_placeholder = classify_cell_format(Some(&format!("__builtin_numFmtId:{id}")));
        assert_eq!(
            from_placeholder, from_string,
            "placeholder should match built-in string for id {id}"
        );
    }
}

#[test]
fn custom_numeric_formats_compute_decimal_counts() {
    let cases: &[(&str, &str)] = &[
        ("0.000", "F3"),
        ("0.0%", "P1"),
        ("0.0E+00", "S1"),
    ];

    for &(fmt, expected) in cases {
        let info = classify_cell_format(Some(fmt));
        assert_eq!(info.cell_format_code, expected, "fmt {fmt}");
        assert!(!info.negative_in_color, "fmt {fmt}");
        assert!(!info.negative_in_parentheses, "fmt {fmt}");
    }
}

#[test]
fn datetime_formats_map_to_cell_d_and_t_codes() {
    let cases: &[(&str, &str)] = &[
        ("m/d/yyyy", "D4"),
        ("h:mm:ss", "D8"),
        ("[h]:mm:ss", "D8"),
        ("mm:ss.0", "D8"),
    ];

    for &(fmt, expected) in cases {
        let info = classify_cell_format(Some(fmt));
        assert_eq!(info.cell_format_code, expected, "fmt {fmt}");
        assert!(!info.negative_in_color, "fmt {fmt}");
        assert!(!info.negative_in_parentheses, "fmt {fmt}");
    }
}

#[test]
fn empty_or_whitespace_formats_are_general() {
    let cases: &[Option<&str>] = &[None, Some(""), Some("   ")];
    for &fmt in cases {
        let info = classify_cell_format(fmt);
        assert_eq!(info.cell_format_code, "G");
    }
}

#[test]
fn unknown_formats_return_n() {
    // Fractions are not part of the fixed/currency/percent/scientific families for CELL("format").
    let info = classify_cell_format(Some("# ?/?"));
    assert_eq!(info.cell_format_code, "N");

    // Non-placeholder, non-numeric literals.
    let info = classify_cell_format(Some("\"hello\""));
    assert_eq!(info.cell_format_code, "N");
}

#[test]
fn negative_parentheses_ignore_layout_fill_operands() {
    // Layout fill token `*X` repeats `X`; `X` is not a literal. Parentheses used only as
    // fill operands must not trigger `CELL("parentheses")` semantics.
    assert!(
        !classify_cell_format(Some("0;0*(0*)")).negative_in_parentheses,
        "fill-token operands should not count as parentheses"
    );

    // Control: explicit parentheses in the negative section should still count.
    assert!(
        classify_cell_format(Some("0;(0)")).negative_in_parentheses,
        "explicit parentheses should count"
    );

    // Underscore layout token `_X` also uses an operand that should be ignored.
    assert!(
        !classify_cell_format(Some("0;0_(0_)")).negative_in_parentheses,
        "underscore-token operands should not count as parentheses"
    );
}

// Ensure the classification struct remains cheap to compare for tests.
#[test]
fn cell_format_classification_is_eq() {
    let a = CellFormatClassification {
        cell_format_code: "G".to_string(),
        negative_in_color: false,
        negative_in_parentheses: false,
    };
    let b = CellFormatClassification {
        cell_format_code: "G".to_string(),
        negative_in_color: false,
        negative_in_parentheses: false,
    };
    assert_eq!(a, b);
}
