#[path = "functions/harness.rs"]
mod harness;

use formula_engine::value::{EntityValue, RecordValue};
use formula_engine::{ErrorKind, Value};

use harness::TestSheet;

#[test]
fn and_rejects_scalar_entity_and_record_like_text() {
    let mut sheet = TestSheet::new();

    sheet.set("A1", Value::Entity(EntityValue::new("Entity")));
    assert_eq!(sheet.eval("=AND(A1)"), Value::Error(ErrorKind::Value));

    sheet.set("A1", Value::Record(RecordValue::new("Record")));
    assert_eq!(sheet.eval("=AND(A1)"), Value::Error(ErrorKind::Value));
}

#[test]
fn and_or_ignore_entity_and_record_in_ranges_like_text() {
    let mut sheet = TestSheet::new();

    sheet.set("A1", Value::Entity(EntityValue::new("Entity")));
    assert_eq!(sheet.eval("=AND(A1:A1)"), Value::Bool(true));
    assert_eq!(sheet.eval("=OR(A1:A1)"), Value::Bool(false));

    sheet.set("A1", Value::Record(RecordValue::new("Record")));
    assert_eq!(sheet.eval("=AND(A1:A1)"), Value::Bool(true));
    assert_eq!(sheet.eval("=OR(A1:A1)"), Value::Bool(false));
}

#[test]
fn not_returns_value_error_for_entity() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Entity(EntityValue::new("Entity")));
    assert_eq!(sheet.eval("=NOT(A1)"), Value::Error(ErrorKind::Value));
}
