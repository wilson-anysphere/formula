use formula_format::{format_value, DateSystem, FormatOptions, Locale, Value};

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

