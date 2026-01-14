#[path = "functions/harness.rs"]
mod harness;

use formula_engine::value::EntityValue;
use formula_engine::Value;

use harness::TestSheet;

#[test]
fn field_access_supports_literal_bracketed_field_names() {
    let mut sheet = TestSheet::new();

    sheet.set(
        "A1",
        Value::Entity(EntityValue::with_properties("Product", [("[Price]", 12.5)])),
    );

    assert_eq!(sheet.eval(r#"=A1.["[Price]"]"#), Value::Number(12.5));
}
