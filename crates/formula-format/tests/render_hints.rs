use formula_format::{
    builtin_format_code, builtin_format_code_with_locale, render_value, AlignmentHint, ColorOverride, DateSystem,
    FormatOptions, LiteralLayoutOp, Locale, Value,
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
        "1.20"
    );
    assert_eq!(
        render_value(Value::Number(1.2), Some("[$€-407]0.00"), &options).text,
        "€1,20"
    );

    assert_eq!(
        render_value(Value::Number(1234.5), Some("[$€-407]#,##0.00"), &options).text,
        "€1.234,50"
    );

    assert_eq!(
        render_value(Value::Number(1234.5), Some("[$€-40C]#,##0.00"), &options).text,
        format!("€1\u{00A0}234,50")
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
fn builtin_accounting_formats_emit_layout_hints_and_can_localize_currency() {
    // Built-in id 44 is the standard accounting format with a currency symbol.
    let fmt_en = builtin_format_code(44).expect("built-in 44");
    let options_en = FormatOptions::default();
    let rendered = render_value(Value::Number(1234.5), Some(fmt_en), &options_en);
    assert!(
        rendered.text.contains('$'),
        "expected built-in accounting format to include a currency symbol: {:?}",
        rendered.text
    );
    let hint = rendered.layout_hint.expect("expected layout hint for accounting format");
    assert!(
        hint.ops
            .iter()
            .any(|op| matches!(op, LiteralLayoutOp::Underscore { .. })),
        "expected underscore ops in accounting format hint: {hint:?}"
    );
    assert!(
        hint.ops.iter().any(|op| matches!(op, LiteralLayoutOp::Fill { .. })),
        "expected fill ops in accounting format hint: {hint:?}"
    );

    // Locale-aware resolver should substitute currency symbol while numeric separators
    // are driven by FormatOptions.locale.
    let fmt_de = builtin_format_code_with_locale(44, Locale::de_de()).expect("built-in 44 de");
    let options_de = FormatOptions {
        locale: Locale::de_de(),
        date_system: DateSystem::Excel1900,
    };
    let rendered = render_value(Value::Number(1234.5), Some(fmt_de.as_ref()), &options_de);
    assert!(
        rendered.text.contains('€'),
        "expected localized accounting currency symbol: {:?}",
        rendered.text
    );
    assert!(
        rendered.text.contains("1.234,50"),
        "expected localized separators in output: {:?}",
        rendered.text
    );
    assert!(rendered.layout_hint.is_some());
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
    assert_eq!(
        render_value(Value::Number(99_999_999_999.0), Some("General"), &options).text,
        "99999999999"
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

    // Literal text must be quoted in Excel format codes.
    let rendered = render_value(Value::Number(1.0), Some("0\"m\""), &options);
    assert_eq!(rendered.text, "1m");
}

#[test]
fn non_finite_numbers_render_as_num_error() {
    let options = FormatOptions::default();
    let rendered = render_value(Value::Number(f64::NAN), None, &options);
    assert_eq!(rendered.text, "#NUM!");
    assert_eq!(rendered.alignment, AlignmentHint::Center);
}
