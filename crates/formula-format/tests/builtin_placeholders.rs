use formula_format::{format_value, FormatOptions, Value};

#[test]
fn builtin_placeholder_renders_percent() {
    let options = FormatOptions::default();
    let formatted = format_value(Value::Number(0.5), Some("__builtin_numFmtId:9"), &options);
    assert!(formatted.text.contains('%'), "expected percent output, got {}", formatted.text);
}

#[test]
fn builtin_placeholder_renders_datetime() {
    let options = FormatOptions::default();
    let formatted = format_value(Value::Number(1.0), Some("__builtin_numFmtId:14"), &options);
    assert_eq!(formatted.text, "1/1/1900");
}

#[test]
fn unknown_builtin_placeholder_falls_back_to_general() {
    let options = FormatOptions::default();
    let formatted = format_value(Value::Number(1234.5), Some("__builtin_numFmtId:999"), &options);
    assert_eq!(formatted.text, "1234.5");
    assert!(
        !formatted.text.contains("__builtin_numFmtId:"),
        "should not render placeholder literal, got {}",
        formatted.text
    );
}

