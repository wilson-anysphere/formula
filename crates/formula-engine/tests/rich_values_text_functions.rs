#[path = "functions/harness.rs"]
mod harness;

use std::collections::HashMap;

use formula_engine::value::{EntityValue, RecordValue};
use formula_engine::Value;

use harness::TestSheet;

#[test]
fn text_and_textjoin_use_display_strings_for_rich_values() {
    let mut sheet = TestSheet::new();

    sheet.set("A1", Value::Entity(EntityValue::new("Apple Inc.")));
    sheet.set(
        "A2",
        Value::Record(RecordValue {
            display: "Fallback".to_string(),
            display_field: Some("Name".to_string()),
            fields: HashMap::from([("Name".to_string(), Value::Text("Apple Inc.".to_string()))]),
        }),
    );

    assert_eq!(
        sheet.eval(r#"=TEXT(A1,"@")"#),
        Value::Text("Apple Inc.".to_string())
    );
    assert_eq!(
        sheet.eval(r#"=TEXTJOIN(",",TRUE,A1,"x")"#),
        Value::Text("Apple Inc.,x".to_string())
    );

    assert_eq!(
        sheet.eval(r#"=TEXT(A2,"@")"#),
        Value::Text("Apple Inc.".to_string())
    );
    assert_eq!(
        sheet.eval(r#"=TEXTJOIN(",",TRUE,A2,"x")"#),
        Value::Text("Apple Inc.,x".to_string())
    );
}
