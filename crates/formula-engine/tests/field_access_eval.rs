#[path = "functions/harness.rs"]
mod harness;

use formula_engine::value::{Array, EntityValue, RecordValue};
use formula_engine::{Engine, ErrorKind, Value};

use harness::TestSheet;

#[test]
fn entity_field_access_returns_property_value() {
    let mut sheet = TestSheet::new();

    sheet.set(
        "A1",
        Value::Entity(EntityValue::with_properties("Product", [("Price", 12.5)])),
    );
    assert_eq!(sheet.eval("=A1.Price"), Value::Number(12.5));
}

#[test]
fn entity_field_access_is_case_insensitive() {
    let mut sheet = TestSheet::new();

    sheet.set(
        "A1",
        Value::Entity(EntityValue::with_properties("Product", [("Price", 12.5)])),
    );
    assert_eq!(sheet.eval("=A1.price"), Value::Number(12.5));
}

#[test]
fn entity_field_access_missing_field_returns_field_error() {
    let mut sheet = TestSheet::new();

    sheet.set(
        "A1",
        Value::Entity(EntityValue::with_properties("Product", [("Price", 12.5)])),
    );
    assert_eq!(sheet.eval("=A1.Missing"), Value::Error(ErrorKind::Field));
}

#[test]
fn field_access_non_rich_base_returns_field_error() {
    let mut sheet = TestSheet::new();

    sheet.set("A1", 1.0);
    assert_eq!(sheet.eval("=A1.Price"), Value::Error(ErrorKind::Field));
}

#[test]
fn record_field_access_returns_field_value() {
    let mut sheet = TestSheet::new();

    sheet.set(
        "A1",
        Value::Record(RecordValue::with_fields_iter("Row", [("Price", 12.5)])),
    );
    assert_eq!(sheet.eval("=A1.Price"), Value::Number(12.5));
}

#[test]
fn field_access_coerces_non_text_field_argument_to_string() {
    let mut sheet = TestSheet::new();

    // `_FIELDACCESS` is a synthetic lowering builtin; make the coercion semantics explicit for
    // direct calls.
    sheet.set(
        "A1",
        Value::Entity(EntityValue::with_properties("Product", [("1", 12.5)])),
    );
    assert_eq!(sheet.eval("=_FIELDACCESS(A1, 1)"), Value::Number(12.5));
}

#[test]
fn field_access_applies_elementwise_over_arrays() {
    let mut engine = Engine::new();

    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Array(Array::new(
                1,
                2,
                vec![
                    Value::Entity(EntityValue::with_properties("A", [("Price", 1.0)])),
                    Value::Entity(EntityValue::with_properties("B", [("Price", 2.0)])),
                ],
            )),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=A1.Price")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(2.0));
}
