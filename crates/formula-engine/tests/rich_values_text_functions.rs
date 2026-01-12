#[path = "functions/harness.rs"]
mod harness;

use formula_engine::value::EntityValue;
use formula_engine::Value;

use harness::TestSheet;

#[test]
fn text_and_textjoin_use_display_strings_for_rich_values() {
    let mut sheet = TestSheet::new();

    sheet.set("A1", Value::Entity(EntityValue::new("Apple Inc.")));

    assert_eq!(
        sheet.eval(r#"=TEXT(A1,"@")"#),
        Value::Text("Apple Inc.".to_string())
    );
    assert_eq!(
        sheet.eval(r#"=TEXTJOIN(",",TRUE,A1,"x")"#),
        Value::Text("Apple Inc.,x".to_string())
    );
}

