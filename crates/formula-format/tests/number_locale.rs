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
        locale::format_number(value, locale::get_locale("es-MX").unwrap()),
        "1,234,567.5"
    );
    assert_eq!(
        locale::format_number(value, locale::get_locale("it-IT").unwrap()),
        "1.234.567,5"
    );
    assert_eq!(
        locale::format_number(value, locale::get_locale("de-CH").unwrap()),
        "1'234'567.5"
    );
    assert_eq!(
        locale::format_number(value, locale::get_locale("fr-CH").unwrap()),
        "1'234'567.5"
    );
    assert_eq!(
        locale::format_number(value, locale::get_locale("it-CH").unwrap()),
        "1'234'567.5"
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
fn does_not_render_negative_zero() {
    assert_eq!(
        locale::format_number(-0.0, locale::get_locale("en-US").unwrap()),
        "0"
    );
    assert_eq!(
        locale::format_number(-0.0, locale::get_locale("fr-FR").unwrap()),
        "0"
    );
}

#[test]
fn get_locale_normalizes_common_locale_id_spellings() {
    assert_eq!(locale::get_locale("en-us").unwrap().id, "en-US");
    assert_eq!(locale::get_locale("en").unwrap().id, "en-US");
    assert_eq!(locale::get_locale("en_gb").unwrap().id, "en-GB");
    assert_eq!(locale::get_locale("en_US.UTF-8").unwrap().id, "en-US");
    assert_eq!(locale::get_locale("de_ch").unwrap().id, "de-CH");
    assert_eq!(locale::get_locale("de_AT").unwrap().id, "de-DE");
    assert_eq!(locale::get_locale("fr_fr").unwrap().id, "fr-FR");
    assert_eq!(locale::get_locale("fr_ch").unwrap().id, "fr-CH");
    assert_eq!(locale::get_locale("fr-CA").unwrap().id, "fr-FR");
    assert_eq!(locale::get_locale("es").unwrap().id, "es-ES");
    assert_eq!(locale::get_locale("es-MX").unwrap().id, "es-MX");
    assert_eq!(locale::get_locale("it").unwrap().id, "it-IT");
    assert_eq!(locale::get_locale("it_CH").unwrap().id, "it-CH");
}

#[test]
fn converts_format_locale_to_number_locale() {
    use formula_format::Locale;

    let fr = locale::number_locale_from_locale(Locale::fr_fr());
    assert_eq!(fr.id, "fr-FR");
    assert_eq!(fr.decimal_separator, ',');
    assert_eq!(fr.thousands_separator, Some('\u{00A0}'));

    let de = locale::number_locale_from_locale(Locale::de_de());
    assert_eq!(de.id, "de-DE");
    assert_eq!(de.decimal_separator, ',');
    assert_eq!(de.thousands_separator, Some('.'));
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
