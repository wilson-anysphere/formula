use formula_format::locale;

#[test]
fn formats_thousands_grouping_for_each_locale() {
    // Pick a value that's exactly representable in binary so `to_string()` is stable.
    let value = 1_234_567.5_f64;
    let nbsp = '\u{00A0}';

    assert_eq!(
        locale::format_number(value, locale::get_locale("en-US").unwrap()),
        "1,234,567.5"
    );
    assert_eq!(
        locale::format_number(value, locale::get_locale("en-GB").unwrap()),
        "1,234,567.5"
    );
    assert_eq!(
        locale::format_number(value, locale::get_locale("de-DE").unwrap()),
        "1.234.567,5"
    );
    assert_eq!(
        locale::format_number(value, locale::get_locale("es-ES").unwrap()),
        "1.234.567,5"
    );
    assert_eq!(
        locale::format_number(value, locale::get_locale("it-IT").unwrap()),
        "1.234.567,5"
    );
    assert_eq!(
        locale::format_number(value, locale::get_locale("fr-FR").unwrap()),
        format!("1{nbsp}234{nbsp}567,5")
    );

    // Also verify that whole numbers omit a decimal separator.
    assert_eq!(
        locale::format_number(1234.0, locale::get_locale("es-ES").unwrap()),
        "1.234"
    );
    assert_eq!(
        locale::format_number(1234.0, locale::get_locale("fr-FR").unwrap()),
        format!("1{nbsp}234")
    );
}

#[test]
fn formats_negative_numbers_for_each_locale() {
    let value = -12_345.5_f64;
    let nbsp = '\u{00A0}';

    assert_eq!(
        locale::format_number(value, locale::get_locale("en-US").unwrap()),
        "-12,345.5"
    );
    assert_eq!(
        locale::format_number(value, locale::get_locale("de-DE").unwrap()),
        "-12.345,5"
    );
    assert_eq!(
        locale::format_number(value, locale::get_locale("fr-FR").unwrap()),
        format!("-12{nbsp}345,5")
    );
}

#[test]
fn localizes_mantissa_decimal_separator_in_scientific_notation() {
    // `f64::to_string()` currently tends to emit fixed-point notation even for values written
    // as `1.23e6`, but the formatter still supports exponent notation if it appears (e.g. on
    // other platforms/toolchains). This test asserts correct behavior for both representations.
    let out = locale::format_number(1.23e6_f64, locale::get_locale("de-DE").unwrap());
    if out.contains('e') || out.contains('E') {
        assert!(
            out.starts_with("1,23"),
            "expected localized mantissa decimal separator, got {out:?}"
        );
    } else {
        assert_eq!(out, "1.230.000");
    }
}

