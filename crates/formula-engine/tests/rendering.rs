use formula_engine::display::format_value_for_display;
use formula_engine::Value;
use formula_format::{AlignmentHint, FormatOptions};

#[test]
fn engine_renders_display_and_alignment() {
    let options = FormatOptions::default();
    let rendered = format_value_for_display(&Value::Number(1.2), Some("0.00"), &options);
    assert_eq!(rendered.text, "1.20");
    assert_eq!(rendered.alignment, AlignmentHint::Right);

    let rendered = format_value_for_display(&Value::Text("abc".to_string()), None, &options);
    assert_eq!(rendered.text, "abc");
    assert_eq!(rendered.alignment, AlignmentHint::Left);

    let rendered = format_value_for_display(&Value::Bool(true), None, &options);
    assert_eq!(rendered.text, "TRUE");
    assert_eq!(rendered.alignment, AlignmentHint::Center);
}
