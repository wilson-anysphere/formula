use formula_format::{
    builtin_format_code, builtin_format_code_with_locale, format_value, DateSystem, FormatOptions, Locale, Value,
};

#[test]
fn builtin_format_ids_cover_ooxml_range() {
    for id in 0u16..=49u16 {
        assert!(
            builtin_format_code(id).is_some(),
            "built-in format {id} missing"
        );
    }
    assert_eq!(builtin_format_code(50), None);
    assert_eq!(builtin_format_code(999), None);
}

#[test]
fn builtin_format_mappings_include_representative_categories() {
    // General / number.
    assert_eq!(builtin_format_code(0), Some("General"));
    assert_eq!(builtin_format_code(1), Some("0"));
    assert_eq!(builtin_format_code(2), Some("0.00"));
    assert_eq!(builtin_format_code(4), Some("#,##0.00"));

    // Percent / scientific / fraction.
    assert_eq!(builtin_format_code(9), Some("0%"));
    assert_eq!(builtin_format_code(11), Some("0.00E+00"));
    assert_eq!(builtin_format_code(12), Some("# ?/?"));

    // Date / time.
    assert_eq!(builtin_format_code(14), Some("m/d/yy"));
    assert_eq!(builtin_format_code(20), Some("h:mm"));
    assert_eq!(builtin_format_code(47), Some("mm:ss.0"));

    // Currency / accounting.
    assert_eq!(builtin_format_code(7), Some("$#,##0.00_);($#,##0.00)"));
    assert_eq!(
        builtin_format_code(44),
        Some(r#"_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)"#)
    );

    // Text.
    assert_eq!(builtin_format_code(49), Some("@"));
}

#[test]
fn builtin_formats_round_trip_through_formatter_under_locales() {
    let en_opts = FormatOptions {
        locale: Locale::en_us(),
        date_system: DateSystem::Excel1900,
    };
    let de_opts = FormatOptions {
        locale: Locale::de_de(),
        date_system: DateSystem::Excel1900,
    };
    let fr_opts = FormatOptions {
        locale: Locale::fr_fr(),
        date_system: DateSystem::Excel1900,
    };

    // General: locale affects decimal separator.
    let general = builtin_format_code(0).unwrap();
    assert_eq!(
        format_value(Value::Number(1234.5), Some(general), &en_opts).text,
        "1234.5"
    );
    assert_eq!(
        format_value(Value::Number(1234.5), Some(general), &de_opts).text,
        "1234,5"
    );
    assert_eq!(
        format_value(Value::Number(1234.5), Some(general), &fr_opts).text,
        "1234,5"
    );

    // Fixed decimals.
    let fixed2 = builtin_format_code(2).unwrap();
    assert_eq!(
        format_value(Value::Number(1234.5), Some(fixed2), &en_opts).text,
        "1234.50"
    );
    assert_eq!(
        format_value(Value::Number(1234.5), Some(fixed2), &de_opts).text,
        "1234,50"
    );
    assert_eq!(
        format_value(Value::Number(1234.5), Some(fixed2), &fr_opts).text,
        "1234,50"
    );

    // Percent scaling.
    let percent = builtin_format_code(9).unwrap();
    assert_eq!(
        format_value(Value::Number(0.256), Some(percent), &en_opts).text,
        "26%"
    );

    // Scientific notation: decimal separator is locale-specific.
    let sci = builtin_format_code(11).unwrap();
    assert_eq!(
        format_value(Value::Number(12345.0), Some(sci), &en_opts).text,
        "1.23E+04"
    );
    assert_eq!(
        format_value(Value::Number(12345.0), Some(sci), &de_opts).text,
        "1,23E+04"
    );
    assert_eq!(
        format_value(Value::Number(12345.0), Some(sci), &fr_opts).text,
        "1,23E+04"
    );

    // Fraction.
    let frac = builtin_format_code(12).unwrap();
    assert_eq!(
        format_value(Value::Number(1.5), Some(frac), &en_opts).text,
        "1 1/2"
    );

    // Date: built-in 14 is locale-variant; use the locale-aware resolver.
    let date_en = builtin_format_code_with_locale(14, Locale::en_us()).unwrap();
    let date_de = builtin_format_code_with_locale(14, Locale::de_de()).unwrap();
    let date_fr = builtin_format_code_with_locale(14, Locale::fr_fr()).unwrap();
    assert_eq!(
        format_value(Value::Number(61.0), Some(date_en.as_ref()), &en_opts).text,
        "3/1/00"
    );
    assert_eq!(
        format_value(Value::Number(61.0), Some(date_de.as_ref()), &de_opts).text,
        "1.3.00"
    );
    assert_eq!(
        format_value(Value::Number(61.0), Some(date_fr.as_ref()), &fr_opts).text,
        "1/3/00"
    );

    // Time with fractional seconds: decimal separator depends on locale.
    let time = builtin_format_code(47).unwrap();
    let serial = 1.234 / 86_400.0;
    assert_eq!(
        format_value(Value::Number(serial), Some(time), &en_opts).text,
        "00:01.2"
    );
    assert_eq!(
        format_value(Value::Number(serial), Some(time), &de_opts).text,
        "00:01,2"
    );
    assert_eq!(
        format_value(Value::Number(serial), Some(time), &fr_opts).text,
        "00:01,2"
    );

    // Currency: locale-aware resolver substitutes the symbol; separators come from options.
    let currency_en = builtin_format_code_with_locale(7, Locale::en_us()).unwrap();
    let currency_de = builtin_format_code_with_locale(7, Locale::de_de()).unwrap();
    let currency_fr = builtin_format_code_with_locale(7, Locale::fr_fr()).unwrap();
    assert_eq!(
        format_value(Value::Number(1234.5), Some(currency_en.as_ref()), &en_opts).text,
        "$1,234.50 "
    );
    assert_eq!(
        format_value(Value::Number(1234.5), Some(currency_de.as_ref()), &de_opts).text,
        "€1.234,50 "
    );
    assert_eq!(
        format_value(Value::Number(1234.5), Some(currency_fr.as_ref()), &fr_opts).text,
        format!("€1\u{00A0}234,50 ")
    );

    // Text.
    let text = builtin_format_code(49).unwrap();
    assert_eq!(
        format_value(Value::Text("hello"), Some(text), &en_opts).text,
        "hello"
    );
}
