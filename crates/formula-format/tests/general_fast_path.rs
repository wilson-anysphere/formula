use formula_format::{format_value, render_value, FormatOptions, Value};

#[test]
fn general_none_matches_explicit_general() {
    let options = FormatOptions::default();

    let a = format_value(Value::Number(1234.5), None, &options);
    let b = format_value(Value::Number(1234.5), Some("General"), &options);
    assert_eq!(a, b);

    // Also assert the full render result matches, to ensure alignment/hints remain consistent.
    let ra = render_value(Value::Number(1234.5), None, &options);
    let rb = render_value(Value::Number(1234.5), Some("General"), &options);
    assert_eq!(ra, rb);
}

