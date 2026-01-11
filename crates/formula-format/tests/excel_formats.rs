use formula_format::{builtin_format_code, format_value, DateSystem, FormatOptions, Locale, Value};

#[test]
fn general_numbers() {
    let options = FormatOptions::default();
    assert_eq!(format_value(Value::Number(1234.5), Some("General"), &options).text, "1234.5");
    assert_eq!(format_value(Value::Number(-2.0), Some("General"), &options).text, "-2");
}

#[test]
fn fixed_decimals_and_grouping() {
    let options = FormatOptions::default();
    assert_eq!(format_value(Value::Number(1.2), Some("0.00"), &options).text, "1.20");
    assert_eq!(format_value(Value::Number(-1.2), Some("0.00"), &options).text, "-1.20");
    assert_eq!(
        format_value(Value::Number(12345.67), Some("#,##0"), &options).text,
        "12,346"
    );
}

#[test]
fn currency_and_percent() {
    let options = FormatOptions::default();
    assert_eq!(
        format_value(Value::Number(1234.5), Some("$#,##0.00"), &options).text,
        "$1,234.50"
    );
    assert_eq!(
        format_value(Value::Number(-1234.5), Some("$#,##0.00"), &options).text,
        "-$1,234.50"
    );

    assert_eq!(format_value(Value::Number(0.1234), Some("0%"), &options).text, "12%");
    assert_eq!(
        format_value(Value::Number(0.1234), Some("0.00%"), &options).text,
        "12.34%"
    );
}

#[test]
fn scientific_notation() {
    let options = FormatOptions::default();
    assert_eq!(
        format_value(Value::Number(12345.0), Some("0.00E+00"), &options).text,
        "1.23E+04"
    );
    assert_eq!(
        format_value(Value::Number(0.0123), Some("0.00E+00"), &options).text,
        "1.23E-02"
    );
}

#[test]
fn fractions() {
    let options = FormatOptions::default();

    assert_eq!(
        format_value(Value::Number(1.5), Some("# ?/?"), &options).text,
        "1 1/2"
    );
    assert_eq!(format_value(Value::Number(0.5), Some("# ?/?"), &options).text, "1/2");
    assert_eq!(format_value(Value::Number(2.0), Some("# ?/?"), &options).text, "2");
}

#[test]
fn dates_1900_system_lotus_bug() {
    let options = FormatOptions {
        locale: Locale::en_us(),
        date_system: DateSystem::Excel1900,
    };

    assert_eq!(
        format_value(Value::Number(1.0), Some("m/d/yyyy"), &options).text,
        "1/1/1900"
    );
    assert_eq!(
        format_value(Value::Number(59.0), Some("m/d/yyyy"), &options).text,
        "2/28/1900"
    );
    // Excel's fictitious 1900-02-29.
    assert_eq!(
        format_value(Value::Number(60.0), Some("m/d/yyyy"), &options).text,
        "2/29/1900"
    );
    assert_eq!(
        format_value(Value::Number(61.0), Some("m/d/yyyy"), &options).text,
        "3/1/1900"
    );

    assert_eq!(
        format_value(Value::Number(1.5), Some("m/d/yyyy h:mm"), &options).text,
        "1/1/1900 12:00"
    );
}

#[test]
fn dates_1904_system_epoch() {
    let options = FormatOptions {
        locale: Locale::en_us(),
        date_system: DateSystem::Excel1904,
    };

    assert_eq!(
        format_value(Value::Number(0.0), Some("m/d/yyyy"), &options).text,
        "1/1/1904"
    );
    assert_eq!(
        format_value(Value::Number(1.0), Some("m/d/yyyy"), &options).text,
        "1/2/1904"
    );
}

#[test]
fn am_pm_time_formatting() {
    let options = FormatOptions {
        locale: Locale::en_us(),
        date_system: DateSystem::Excel1900,
    };

    assert_eq!(
        format_value(Value::Number(1.0), Some("h:mm AM/PM"), &options).text,
        "12:00 AM"
    );
    assert_eq!(
        format_value(Value::Number(1.5), Some("h:mm AM/PM"), &options).text,
        "12:00 PM"
    );
    assert_eq!(
        format_value(Value::Number(1.75), Some("h:mm:ss AM/PM"), &options).text,
        "6:00:00 PM"
    );

    assert_eq!(
        format_value(Value::Number(1.0), Some("h:mm A/P"), &options).text,
        "12:00 A"
    );
    assert_eq!(
        format_value(Value::Number(1.5), Some("h:mm A/P"), &options).text,
        "12:00 P"
    );
}

#[test]
fn fractional_seconds_time_formatting() {
    let options = FormatOptions {
        locale: Locale::en_us(),
        date_system: DateSystem::Excel1900,
    };

    assert_eq!(
        format_value(Value::Number(0.0), Some("mm:ss.0"), &options).text,
        "00:00.0"
    );

    let serial = 1.234 / 86_400.0;
    assert_eq!(
        format_value(Value::Number(serial), Some("mm:ss.0"), &options).text,
        "00:01.2"
    );

    let serial = 59.96 / 86_400.0;
    assert_eq!(
        format_value(Value::Number(serial), Some("mm:ss.0"), &options).text,
        "01:00.0"
    );
}

#[test]
fn locale_separators() {
    let options = FormatOptions {
        locale: Locale::de_de(),
        date_system: DateSystem::Excel1900,
    };
    assert_eq!(
        format_value(Value::Number(1234.5), Some("#,##0.00"), &options).text,
        "1.234,50"
    );
    assert_eq!(
        format_value(Value::Number(1.0), Some("m/d/yyyy"), &options).text,
        "1.1.1900"
    );

    let serial = 1.234 / 86_400.0;
    assert_eq!(
        format_value(Value::Number(serial), Some("mm:ss.0"), &options).text,
        "00:01,2"
    );
}

#[test]
fn conditional_sections_and_text() {
    let options = FormatOptions::default();
    let code = r#"0.0;[Red]-0.0;"zero";@"#;

    assert_eq!(format_value(Value::Number(1.0), Some(code), &options).text, "1.0");
    assert_eq!(format_value(Value::Number(-1.0), Some(code), &options).text, "-1.0");
    assert_eq!(format_value(Value::Number(0.0), Some(code), &options).text, "zero");
    assert_eq!(
        format_value(Value::Text("hello"), Some(code), &options).text,
        "hello"
    );
}

#[test]
fn error_values_align_center_like_excel() {
    let options = FormatOptions::default();
    let rendered = format_value(Value::Error("#DIV/0!"), None, &options);
    assert_eq!(rendered.text, "#DIV/0!");
    assert_eq!(rendered.alignment, formula_format::AlignmentHint::Center);
}

#[test]
fn bracket_currency_tokens_render_currency_symbol() {
    let options = FormatOptions::default();
    assert_eq!(
        format_value(Value::Number(1234.5), Some("[$$-409]#,##0.00"), &options).text,
        "$1,234.50"
    );
    assert_eq!(
        format_value(
            Value::Number(-1234.5),
            Some("[$$-409]#,##0.00;([$$-409]#,##0.00)"),
            &options
        )
        .text,
        "($1,234.50)"
    );
}

#[test]
fn text_at_placeholder_formats_numbers_as_text_and_aligns_left() {
    let options = FormatOptions::default();
    let rendered = format_value(Value::Number(123.0), Some("@"), &options);
    assert_eq!(rendered.text, "123");
    assert_eq!(rendered.alignment, formula_format::AlignmentHint::Left);

    let rendered = format_value(Value::Number(-12.5), Some("\"Value: \"@"), &options);
    assert_eq!(rendered.text, "Value: -12.5");
    assert_eq!(rendered.alignment, formula_format::AlignmentHint::Left);

    let rendered = format_value(Value::Text("hello"), Some("\"Value: \"@"), &options);
    assert_eq!(rendered.text, "Value: hello");
}

#[test]
fn builtin_currency_and_accounting_formats_smoke() {
    let options = FormatOptions::default();

    // Built-in currency format (2 decimals).
    let currency = builtin_format_code(7).expect("missing built-in currency id 7");
    let rendered = format_value(Value::Number(1234.5), Some(currency), &options).text;
    assert!(
        rendered.contains("$1,234.50"),
        "expected currency output, got {rendered:?}"
    );
    let rendered_neg = format_value(Value::Number(-1234.5), Some(currency), &options).text;
    assert!(
        rendered_neg.contains("($1,234.50)"),
        "expected currency negative output, got {rendered_neg:?}"
    );

    // Built-in accounting format (2 decimals).
    let accounting = builtin_format_code(44).expect("missing built-in accounting id 44");
    let rendered = format_value(Value::Number(1234.5), Some(accounting), &options).text;
    assert!(
        rendered.contains("$") && rendered.contains("1,234.50"),
        "expected accounting output, got {rendered:?}"
    );

    let rendered_zero = format_value(Value::Number(0.0), Some(accounting), &options).text;
    assert!(
        !rendered_zero.trim().is_empty() && rendered_zero.contains('-'),
        "expected accounting zero section output, got {rendered_zero:?}"
    );
}
