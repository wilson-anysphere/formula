use formula_format::{
    render_value, AlignmentHint, ColorOverride, DateSystem, FormatOptions, LiteralLayoutOp, Locale, Value,
};

#[test]
fn color_tokens_are_exposed_as_hints_and_not_rendered() {
    let options = FormatOptions::default();

    // Two-section format; negative section has a color and no explicit '-' so Excel shows the
    // absolute value.
    let rendered = render_value(Value::Number(-12.0), Some("0;[Red]0"), &options);
    assert_eq!(rendered.text, "12");
    assert_eq!(rendered.color, Some(ColorOverride::Argb(0xFFFF0000)));
    assert_eq!(rendered.alignment, AlignmentHint::Right);

    let rendered = render_value(Value::Number(-12.0), Some("0;[Color10]0"), &options);
    assert_eq!(rendered.text, "12");
    assert_eq!(rendered.color, Some(ColorOverride::Indexed(10)));
}

#[test]
fn bracket_currency_tokens_resolve_default_symbol_for_locale() {
    let options = FormatOptions::default();
    assert_eq!(
        render_value(Value::Number(1.2), Some("[$-409]0.00"), &options).text,
        "$1.20"
    );
    assert_eq!(
        render_value(Value::Number(1.2), Some("[$€-407]0.00"), &options).text,
        "€1.20"
    );
}

#[test]
fn accounting_underscores_and_fill_emit_layout_hints() {
    let options = FormatOptions::default();

    let rendered = render_value(Value::Number(1234.0), Some("#,##0_);(#,##0)"), &options);
    assert_eq!(rendered.text, "1,234 ");
    let hint = rendered.layout_hint.expect("layout hint");
    assert_eq!(
        hint.ops,
        vec![LiteralLayoutOp::Underscore {
            byte_index: 5,
            width_of: ')'
        }]
    );

    let rendered = render_value(Value::Number(5.0), Some("*-0"), &options);
    assert_eq!(rendered.text, "5");
    let hint = rendered.layout_hint.expect("layout hint");
    assert_eq!(
        hint.ops,
        vec![LiteralLayoutOp::Fill {
            byte_index: 0,
            fill_with: '-'
        }]
    );
}

#[test]
fn general_format_matches_excel_like_rules() {
    let options = FormatOptions::default();

    assert_eq!(
        render_value(Value::Number(-0.0), Some("General"), &options).text,
        "0"
    );
    assert_eq!(
        render_value(Value::Number(1.234567890123456), Some("General"), &options).text,
        "1.23456789012346"
    );
    assert_eq!(
        render_value(Value::Number(1e-10), Some("General"), &options).text,
        "1E-10"
    );
    assert_eq!(
        render_value(Value::Number(1e11), Some("General"), &options).text,
        "1E+11"
    );
}

#[test]
fn datetime_month_only_formats_are_detected_and_bracket_tokens_ignored() {
    let options = FormatOptions {
        locale: Locale::en_us(),
        date_system: DateSystem::Excel1900,
    };

    // Serial 1.0 is 1900-01-01 in the 1900 system.
    let rendered = render_value(Value::Number(1.0), Some("mm"), &options);
    assert_eq!(rendered.text, "01");

    let rendered = render_value(Value::Number(1.0), Some("[Red]mm"), &options);
    assert_eq!(rendered.text, "01");
    assert_eq!(rendered.color, Some(ColorOverride::Argb(0xFFFF0000)));

    // Numeric formats containing `m` should remain numeric when placeholders are present.
    let rendered = render_value(Value::Number(1.0), Some("0m"), &options);
    assert_eq!(rendered.text, "1m");
}

#[test]
fn non_finite_numbers_render_as_num_error() {
    let options = FormatOptions::default();
    let rendered = render_value(Value::Number(f64::NAN), None, &options);
    assert_eq!(rendered.text, "#NUM!");
    assert_eq!(rendered.alignment, AlignmentHint::Center);
}

